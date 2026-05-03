use super::utils::find_stream_case;
use super::vba::{normalize_vba_source_text, parse_vba_project_text, vba_decompress};
use super::{hex, ParseError, ParserConfig};
use crate::ole::Cfb;
use crate::zip_handler::PackageReader;
use docir_core::ir::IRNode;
use docir_core::ir::IrBuilder;
use docir_core::security::analyze_vba_source;
use docir_core::security::{MacroExtractionState, MacroModuleType, MacroProject};
use docir_core::types::SourceSpan;
use docir_core::visitor::IrStore;

pub(super) struct VbaScanner<'a> {
    config: &'a ParserConfig,
}

impl<'a> VbaScanner<'a> {
    pub(super) fn new(config: &'a ParserConfig) -> Self {
        Self { config }
    }

    pub(super) fn scan_zip_vba_projects(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        let mut builder = IrBuilder::new(store);
        let vba_paths = [
            "word/vbaProject.bin",
            "xl/vbaProject.bin",
            "ppt/vbaProject.bin",
        ];
        for vba_path in &vba_paths {
            if zip.contains(vba_path) {
                let (mut macro_project, modules) = self.detect_macro_project(zip, vba_path)?;
                for module in modules {
                    let id = module.id;
                    builder.insert(IRNode::MacroModule(module));
                    macro_project.modules.push(id);
                }
                builder.insert(IRNode::MacroProject(macro_project));
            }
        }
        Ok(())
    }

    pub(super) fn scan_cfb_vba_projects(
        &self,
        cfb: &Cfb,
        streams: &[String],
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        let mut builder = IrBuilder::new(store);
        let mut seen_roots = std::collections::HashSet::new();

        for project_stream in streams.iter().filter(|path| {
            let upper = path.to_ascii_uppercase();
            upper == "PROJECT" || upper.ends_with("/PROJECT")
        }) {
            let storage_root = project_stream
                .rsplit_once('/')
                .map(|(parent, _)| parent.to_string())
                .unwrap_or_default();
            if !seen_roots.insert(storage_root.clone()) {
                continue;
            }

            let container_label = if storage_root.is_empty() {
                "cfb:/PROJECT".to_string()
            } else {
                format!("cfb:/{storage_root}")
            };
            let (mut macro_project, modules) =
                self.detect_macro_project_in_cfb(cfb, &container_label, &storage_root)?;
            for module in modules {
                let id = module.id;
                builder.insert(IRNode::MacroModule(module));
                macro_project.modules.push(id);
            }
            builder.insert(IRNode::MacroProject(macro_project));
        }
        Ok(())
    }

    fn detect_macro_project(
        &self,
        zip: &mut impl PackageReader,
        path: &str,
    ) -> Result<(MacroProject, Vec<docir_core::security::MacroModule>), ParseError> {
        let data = zip.read_file(path)?;

        let cfb = crate::ole::Cfb::parse(data)?;
        self.detect_macro_project_in_cfb(&cfb, path, "VBA")
    }

    fn detect_macro_project_in_cfb(
        &self,
        cfb: &Cfb,
        container_path: &str,
        storage_root: &str,
    ) -> Result<(MacroProject, Vec<docir_core::security::MacroModule>), ParseError> {
        let streams = cfb.list_streams();
        let project_stream_name = if storage_root.is_empty() {
            "PROJECT".to_string()
        } else {
            format!("{storage_root}/PROJECT")
        };

        let mut project = MacroProject::new();
        project.container_path = Some(container_path.to_string());
        project.storage_root = Some(storage_root.to_string());
        project.span = Some(SourceSpan::new(container_path));

        let project_stream =
            find_stream_case(&streams, &project_stream_name).and_then(|p| cfb.read_stream(p));
        let project_text = project_stream
            .as_ref()
            .map(|data| String::from_utf8_lossy(data).to_string())
            .unwrap_or_default();

        let (project_name, module_defs, references, is_protected) =
            parse_vba_project_text(&project_text);
        project.name = project_name;
        project.references = references;
        project.is_protected = is_protected;

        let (modules_out, auto_exec) =
            self.extract_vba_modules(cfb, &streams, container_path, storage_root, &module_defs);

        if !auto_exec.is_empty() {
            project.has_auto_exec = true;
            project.auto_exec_procedures = auto_exec;
        }

        Ok((project, modules_out))
    }

    fn extract_vba_modules(
        &self,
        cfb: &Cfb,
        streams: &[String],
        container_path: &str,
        storage_root: &str,
        module_defs: &[(String, MacroModuleType)],
    ) -> (Vec<docir_core::security::MacroModule>, Vec<String>) {
        let mut auto_exec = Vec::new();
        let mut modules_out = Vec::new();
        for (module_name, module_type) in module_defs {
            let stream_path = if storage_root.is_empty() {
                module_name.clone()
            } else {
                format!("{storage_root}/{module_name}")
            };
            let mut module =
                docir_core::security::MacroModule::new(module_name.clone(), *module_type);
            module.stream_name = Some(module_name.clone());
            module.stream_path = Some(stream_path.clone());
            module.span = Some(SourceSpan::new(container_path));

            let raw_stream =
                find_stream_case(streams, &stream_path).and_then(|p| cfb.read_stream(p));
            match raw_stream {
                Some(raw) => {
                    module.compressed_size = Some(raw.len() as u64);
                    let Some(src) = vba_decompress(&raw) else {
                        module.extraction_state = MacroExtractionState::DecodeFailed;
                        module
                            .extraction_errors
                            .push(format!("Failed to decompress {}", stream_path));
                        modules_out.push(module);
                        continue;
                    };

                    module.decompressed_size = Some(src.len() as u64);
                    module.extraction_state = MacroExtractionState::Extracted;
                    module.source_encoding = Some("utf-8-lossy-normalized".to_string());
                    let source = normalize_vba_source_text(&src);
                    let analysis = analyze_vba_source(&source);
                    auto_exec.extend(analysis.auto_exec_procedures.clone());
                    module.procedures = analysis.procedures;
                    module.suspicious_calls = analysis.suspicious_calls;

                    if self.config.extract_macro_source {
                        module.source_code = Some(source);
                    }
                    if self.config.compute_hashes {
                        use sha2::Digest;
                        let mut hasher = sha2::Sha256::new();
                        hasher.update(&src);
                        module.source_hash = Some(hex::encode(hasher.finalize()));
                    }
                }
                None => {
                    module.extraction_state = MacroExtractionState::MissingStream;
                    module
                        .extraction_errors
                        .push(format!("Missing stream {}", stream_path));
                }
            };

            modules_out.push(module);
        }
        (modules_out, auto_exec)
    }
}
