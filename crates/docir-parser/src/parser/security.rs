use super::utils::find_stream_case;
use super::vba::{parse_vba_project_text, vba_decompress};
use super::{hex, ParseError, ParserConfig};
use crate::ooxml::part_utils::get_rels_path;
use crate::ooxml::relationships::{rel_type, Relationships, TargetMode};
use crate::zip_handler::SecureZipReader;
use docir_core::ir::IRNode;
use docir_core::security::{ExternalRefType, ExternalReference, MacroProject, OleObject};
use docir_core::types::SourceSpan;
use docir_core::visitor::IrStore;
use std::collections::HashSet;
use std::io::{Read, Seek};

pub(super) struct SecurityScanner<'a> {
    config: &'a ParserConfig,
}

impl<'a> SecurityScanner<'a> {
    pub(super) fn new(config: &'a ParserConfig) -> Self {
        Self { config }
    }

    pub(super) fn scan_activex_controls<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        let mut activex_bin_seen: HashSet<String> = HashSet::new();
        let activex_paths: Vec<String> = zip
            .list_prefix("word/activeX/")
            .into_iter()
            .chain(zip.list_prefix("xl/activeX/"))
            .chain(zip.list_prefix("ppt/activeX/"))
            .filter(|p| p.ends_with(".xml"))
            .map(|s| s.to_string())
            .collect();
        for path in activex_paths {
            let xml = zip.read_file_string(&path)?;
            if let Some(mut control) = super::parse_activex_xml(&xml, &path) {
                control.span = Some(SourceSpan::new(&path));
                store.insert(IRNode::ActiveXControl(control));
            }

            let rels_path = get_rels_path(&path);
            if zip.contains(&rels_path) {
                if let Ok(rels_xml) = zip.read_file_string(&rels_path) {
                    if let Ok(rels) = Relationships::parse(&rels_xml) {
                        for rel in rels.by_id.values() {
                            if !rel.target.ends_with(".bin")
                                && !rel.rel_type.contains("activeXControlBinary")
                            {
                                continue;
                            }
                            let bin_path = Relationships::resolve_target(&path, &rel.target);
                            if activex_bin_seen.insert(bin_path.clone()) && zip.contains(&bin_path)
                            {
                                let ole_object = self.detect_ole_object(zip, &bin_path)?;
                                store.insert(IRNode::OleObject(ole_object));
                            }
                        }
                    }
                }
            }
        }
        Ok(())
    }

    pub(super) fn scan_word_external_relationships<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        let rel_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.starts_with("word/") && p.ends_with(".rels"))
            .map(|s| s.to_string())
            .collect();
        for rel_path in rel_paths {
            let rels_xml = zip.read_file_string(&rel_path)?;
            let rels = Relationships::parse(&rels_xml)?;
            for rel in rels.by_id.values() {
                if rel.target_mode == TargetMode::External {
                    let ref_type = match rel.rel_type.as_str() {
                        rel_type::HYPERLINK => ExternalRefType::Hyperlink,
                        rel_type::IMAGE => ExternalRefType::Image,
                        rel_type::OLE_OBJECT => ExternalRefType::OleLink,
                        rel_type::ATTACHED_TEMPLATE => ExternalRefType::AttachedTemplate,
                        _ => ExternalRefType::Other,
                    };
                    let mut ext_ref = ExternalReference::new(ref_type, &rel.target);
                    ext_ref.relationship_id = Some(rel.id.clone());
                    ext_ref.relationship_type = Some(rel.rel_type.clone());
                    ext_ref.span = Some(SourceSpan::new(&rel_path));
                    store.insert(IRNode::ExternalReference(ext_ref));
                }
            }
        }
        Ok(())
    }

    pub(super) fn detect_macro_project<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        path: &str,
    ) -> Result<(MacroProject, Vec<docir_core::security::MacroModule>), ParseError> {
        let data = zip.read_file(path)?;

        let mut project = MacroProject::new();
        project.span = Some(SourceSpan::new(path));

        let cfb = crate::ole::Cfb::parse(data)?;
        let streams = cfb.list_streams();

        let project_stream =
            find_stream_case(&streams, "VBA/PROJECT").and_then(|p| cfb.read_stream(p));
        let project_text = project_stream
            .as_ref()
            .map(|data| String::from_utf8_lossy(data).to_string())
            .unwrap_or_default();

        let (project_name, module_defs, references, is_protected) =
            parse_vba_project_text(&project_text);
        project.name = project_name;
        project.references = references;
        project.is_protected = is_protected;

        let mut auto_exec = Vec::new();
        let mut modules_out = Vec::new();
        for (module_name, module_type) in module_defs {
            let stream_path = format!("VBA/{module_name}");
            let data = find_stream_case(&streams, &stream_path)
                .and_then(|p| cfb.read_stream(p))
                .and_then(|raw| vba_decompress(&raw));

            let mut module =
                docir_core::security::MacroModule::new(module_name.clone(), module_type);
            module.span = Some(SourceSpan::new(path));

            if let Some(src) = data {
                let source = String::from_utf8_lossy(&src).to_string();
                let analysis = docir_core::security::analyze_vba_source(&source);
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

            modules_out.push(module);
        }

        if !auto_exec.is_empty() {
            project.has_auto_exec = true;
            project.auto_exec_procedures = auto_exec;
        }

        Ok((project, modules_out))
    }

    pub(super) fn detect_ole_object<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        path: &str,
    ) -> Result<OleObject, ParseError> {
        let data = zip.read_file(path)?;

        let mut ole = OleObject::new();
        ole.span = Some(SourceSpan::new(path));
        ole.size_bytes = data.len() as u64;

        if self.config.compute_hashes {
            use sha2::{Digest, Sha256};
            let mut hasher = Sha256::new();
            hasher.update(&data);
            let hash = hasher.finalize();
            ole.data_hash = Some(hex::encode(hash));
        }

        Ok(ole)
    }
}
