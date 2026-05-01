use crate::ParsedDocument;
use docir_core::ir::IRNode;
use docir_core::security::{
    MacroExtractionState, MacroModule, MacroModuleType, MacroProject, SuspiciousCall,
};
use serde::Serialize;

/// High-level recognition status for a VBA extraction attempt.
#[derive(Debug, Clone, Copy, Serialize, PartialEq, Eq)]
pub enum VbaRecognitionStatus {
    Absent,
    Recognized,
    Extracted,
    Partial,
    Error,
}

/// Structured VBA recognition report for downstream tools.
#[derive(Debug, Clone, Serialize)]
pub struct VbaRecognitionReport {
    pub document_format: String,
    pub status: VbaRecognitionStatus,
    pub projects: Vec<VbaProjectReport>,
}

/// Report for a single VBA project.
#[derive(Debug, Clone, Serialize)]
pub struct VbaProjectReport {
    pub node_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub project_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub container_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub storage_root: Option<String>,
    pub is_protected: bool,
    pub has_auto_exec: bool,
    #[serde(default)]
    pub auto_exec_procedures: Vec<String>,
    #[serde(default)]
    pub references: Vec<String>,
    pub status: VbaRecognitionStatus,
    pub modules: Vec<VbaModuleReport>,
}

/// Report for a single VBA module.
#[derive(Debug, Clone, Serialize)]
pub struct VbaModuleReport {
    pub node_id: String,
    pub name: String,
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stream_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stream_path: Option<String>,
    pub status: VbaRecognitionStatus,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_encoding: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_hash: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub compressed_size: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub decompressed_size: Option<u64>,
    #[serde(default)]
    pub procedures: Vec<String>,
    #[serde(default)]
    pub suspicious_calls: Vec<String>,
    #[serde(default)]
    pub extraction_errors: Vec<String>,
}

impl VbaRecognitionReport {
    /// Builds a recognition report from the parsed IR.
    pub fn from_parsed(parsed: &ParsedDocument, include_source: bool) -> Self {
        let mut projects = Vec::new();

        for node in parsed.store().values() {
            let IRNode::MacroProject(project) = node else {
                continue;
            };
            projects.push(VbaProjectReport::from_project(
                parsed,
                project,
                include_source,
            ));
        }

        let status = aggregate_project_status(&projects);
        Self {
            document_format: parsed.format().extension().to_string(),
            status,
            projects,
        }
    }
}

impl VbaProjectReport {
    fn from_project(parsed: &ParsedDocument, project: &MacroProject, include_source: bool) -> Self {
        let modules: Vec<_> = project
            .modules
            .iter()
            .filter_map(|id| match parsed.store().get(*id) {
                Some(IRNode::MacroModule(module)) => {
                    Some(VbaModuleReport::from_module(module, include_source))
                }
                _ => None,
            })
            .collect();

        let status = aggregate_module_status(&modules);
        Self {
            node_id: project.id.to_string(),
            project_name: project.name.clone(),
            container_path: project.container_path.clone(),
            storage_root: project.storage_root.clone(),
            is_protected: project.is_protected,
            has_auto_exec: project.has_auto_exec,
            auto_exec_procedures: project.auto_exec_procedures.clone(),
            references: project.references.iter().map(|r| r.name.clone()).collect(),
            status,
            modules,
        }
    }
}

impl VbaModuleReport {
    fn from_module(module: &MacroModule, include_source: bool) -> Self {
        let source_text = include_source.then(|| module.source_code.clone()).flatten();
        Self {
            node_id: module.id.to_string(),
            name: module.name.clone(),
            kind: module_kind_name(module.module_type).to_string(),
            stream_name: module.stream_name.clone(),
            stream_path: module.stream_path.clone(),
            status: map_module_state(
                module.extraction_state,
                &module.extraction_errors,
                include_source,
            ),
            source_text,
            source_encoding: module.source_encoding.clone(),
            source_hash: module.source_hash.clone(),
            compressed_size: module.compressed_size,
            decompressed_size: module.decompressed_size,
            procedures: module.procedures.clone(),
            suspicious_calls: module
                .suspicious_calls
                .iter()
                .map(format_suspicious_call)
                .collect(),
            extraction_errors: module.extraction_errors.clone(),
        }
    }
}

fn aggregate_project_status(projects: &[VbaProjectReport]) -> VbaRecognitionStatus {
    if projects.is_empty() {
        return VbaRecognitionStatus::Absent;
    }
    if projects
        .iter()
        .any(|p| p.status == VbaRecognitionStatus::Partial)
    {
        return VbaRecognitionStatus::Partial;
    }
    if projects
        .iter()
        .any(|p| p.status == VbaRecognitionStatus::Error)
    {
        return VbaRecognitionStatus::Error;
    }
    if projects
        .iter()
        .all(|p| p.status == VbaRecognitionStatus::Extracted)
    {
        return VbaRecognitionStatus::Extracted;
    }
    VbaRecognitionStatus::Recognized
}

fn aggregate_module_status(modules: &[VbaModuleReport]) -> VbaRecognitionStatus {
    if modules.is_empty() {
        return VbaRecognitionStatus::Recognized;
    }
    if modules
        .iter()
        .any(|m| m.status == VbaRecognitionStatus::Error)
    {
        return VbaRecognitionStatus::Error;
    }
    if modules
        .iter()
        .any(|m| m.status == VbaRecognitionStatus::Partial)
    {
        return VbaRecognitionStatus::Partial;
    }
    if modules
        .iter()
        .all(|m| m.status == VbaRecognitionStatus::Extracted)
    {
        return VbaRecognitionStatus::Extracted;
    }
    VbaRecognitionStatus::Recognized
}

fn map_module_state(
    state: MacroExtractionState,
    errors: &[String],
    include_source: bool,
) -> VbaRecognitionStatus {
    match state {
        MacroExtractionState::Extracted if include_source => VbaRecognitionStatus::Extracted,
        MacroExtractionState::Extracted => VbaRecognitionStatus::Recognized,
        MacroExtractionState::Partial => VbaRecognitionStatus::Partial,
        MacroExtractionState::DecodeFailed => VbaRecognitionStatus::Error,
        MacroExtractionState::MissingStream if errors.is_empty() => {
            VbaRecognitionStatus::Recognized
        }
        MacroExtractionState::MissingStream => VbaRecognitionStatus::Partial,
        MacroExtractionState::NotRequested => VbaRecognitionStatus::Recognized,
    }
}

fn module_kind_name(kind: MacroModuleType) -> &'static str {
    match kind {
        MacroModuleType::Standard => "standard",
        MacroModuleType::Class => "class",
        MacroModuleType::UserForm => "form",
        MacroModuleType::Document => "document",
    }
}

fn format_suspicious_call(call: &SuspiciousCall) -> String {
    match call.line {
        Some(line) => format!("{}@{}", call.name, line),
        None => call.name.clone(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_support::build_test_cfb;
    use crate::{DocirApp, ParserConfig};
    use docir_core::ir::{Document, IRNode};
    use docir_core::visitor::IrStore;
    use docir_core::DocumentFormat;

    #[test]
    fn report_marks_absent_when_no_projects_exist() {
        let mut store = IrStore::new();
        let doc = Document::new(DocumentFormat::WordProcessing);
        let root = doc.id;
        store.insert(IRNode::Document(doc));
        let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
            root_id: root,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        });

        let report = VbaRecognitionReport::from_parsed(&parsed, false);
        assert_eq!(report.status, VbaRecognitionStatus::Absent);
        assert!(report.projects.is_empty());
    }

    #[test]
    fn report_extracts_legacy_cfb_vba_project() {
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="LegacyProject"
Module=Module1/Module1
Document=ThisDocument/&H00000000
DPB="AAAA"
"#,
            ),
            ("VBA/Module1", b"Sub AutoOpen()\nShell \"cmd\"\nEnd Sub\n"),
            ("VBA/ThisDocument", b"Sub Document_Open()\nEnd Sub\n"),
        ]);
        let config = ParserConfig {
            extract_macro_source: true,
            ..Default::default()
        };
        let app = DocirApp::new(config);
        let parsed = app.parse_bytes(&bytes).expect("parse legacy cfb");

        let report = app.build_vba_recognition(&parsed, true);
        assert_eq!(report.status, VbaRecognitionStatus::Extracted);
        assert_eq!(report.projects.len(), 1);
        assert_eq!(
            report.projects[0].project_name.as_deref(),
            Some("LegacyProject")
        );
        assert_eq!(report.projects[0].storage_root.as_deref(), Some("VBA"));
        assert!(report.projects[0].is_protected);
        assert!(report.projects[0].has_auto_exec);
        assert!(report.projects[0]
            .auto_exec_procedures
            .iter()
            .any(|name| name.eq_ignore_ascii_case("AutoOpen")));
        assert_eq!(report.projects[0].modules.len(), 2);
        assert!(report.projects[0]
            .modules
            .iter()
            .any(
                |module| module.stream_path.as_deref() == Some("VBA/Module1")
                    && module
                        .source_text
                        .as_deref()
                        .unwrap_or_default()
                        .contains("AutoOpen")
            ));
    }

    #[test]
    fn report_extracts_legacy_xls_vba_project() {
        let bytes = build_test_cfb(&[
            ("Workbook", b"wb"),
            (
                "VBA/PROJECT",
                br#"Name="WorkbookMacros"
Module=Module1/Module1
Document=ThisWorkbook/&H00000000
"#,
            ),
            ("VBA/Module1", b"Sub Auto_Open()\nEnd Sub\n"),
            (
                "VBA/ThisWorkbook",
                b"Private Sub Workbook_Open()\nEnd Sub\n",
            ),
        ]);
        let config = ParserConfig {
            extract_macro_source: true,
            ..Default::default()
        };
        let app = DocirApp::new(config);
        let parsed = app.parse_bytes(&bytes).expect("parse legacy xls");

        let report = app.build_vba_recognition(&parsed, true);
        assert_eq!(report.status, VbaRecognitionStatus::Extracted);
        assert_eq!(report.projects.len(), 1);
        assert_eq!(
            report.projects[0].project_name.as_deref(),
            Some("WorkbookMacros")
        );
        assert_eq!(report.projects[0].storage_root.as_deref(), Some("VBA"));
        assert!(report.projects[0].has_auto_exec);
        assert!(report.projects[0]
            .auto_exec_procedures
            .iter()
            .any(|name| name.eq_ignore_ascii_case("Auto_Open")
                || name.eq_ignore_ascii_case("Workbook_Open")));
        assert_eq!(report.projects[0].modules.len(), 2);
        assert!(report.projects[0]
            .modules
            .iter()
            .any(
                |module| module.stream_path.as_deref() == Some("VBA/ThisWorkbook")
                    && module
                        .source_text
                        .as_deref()
                        .unwrap_or_default()
                        .contains("Workbook_Open")
            ));
    }

    #[test]
    fn report_extracts_legacy_ppt_vba_project() {
        let bytes = build_test_cfb(&[
            ("PowerPoint Document", b"ppt"),
            ("Current User", b"user"),
            (
                "VBA/PROJECT",
                br#"Name="PresentationMacros"
Module=Module1/Module1
Document=ThisPresentation/&H00000000
"#,
            ),
            ("VBA/Module1", b"Sub Auto_Open()\nEnd Sub\n"),
            (
                "VBA/ThisPresentation",
                b"Private Sub Presentation_Open()\nEnd Sub\n",
            ),
        ]);
        let config = ParserConfig {
            extract_macro_source: true,
            ..Default::default()
        };
        let app = DocirApp::new(config);
        let parsed = app.parse_bytes(&bytes).expect("parse legacy ppt");

        let report = app.build_vba_recognition(&parsed, true);
        assert_eq!(report.status, VbaRecognitionStatus::Extracted);
        assert_eq!(report.projects.len(), 1);
        assert_eq!(
            report.projects[0].project_name.as_deref(),
            Some("PresentationMacros")
        );
        assert_eq!(report.projects[0].storage_root.as_deref(), Some("VBA"));
        assert!(report.projects[0].has_auto_exec);
        assert!(report.projects[0]
            .auto_exec_procedures
            .iter()
            .any(|name| name.eq_ignore_ascii_case("Auto_Open")
                || name.eq_ignore_ascii_case("Presentation_Open")));
        assert_eq!(report.projects[0].modules.len(), 2);
        assert!(report.projects[0]
            .modules
            .iter()
            .any(
                |module| module.stream_path.as_deref() == Some("VBA/ThisPresentation")
                    && module
                        .source_text
                        .as_deref()
                        .unwrap_or_default()
                        .contains("Presentation_Open")
            ));
    }

    #[test]
    fn report_marks_legacy_project_protected_and_keeps_references() {
        let bytes = build_test_cfb(&[
            ("Workbook", b"wb"),
            (
                "VBA/PROJECT",
                br#"Name="ProtectedWorkbook"
Module=Module1/Module1
Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation
DPB="AAAA"
"#,
            ),
            ("VBA/Module1", b"Sub Auto_Open()\nEnd Sub\n"),
        ]);
        let config = ParserConfig {
            extract_macro_source: true,
            ..Default::default()
        };
        let app = DocirApp::new(config);
        let parsed = app.parse_bytes(&bytes).expect("parse protected legacy xls");

        let report = app.build_vba_recognition(&parsed, true);
        assert_eq!(report.status, VbaRecognitionStatus::Extracted);
        assert_eq!(report.projects.len(), 1);
        assert!(report.projects[0].is_protected);
        assert_eq!(report.projects[0].references.len(), 1);
        assert!(report.projects[0].references[0].contains("OLE Automation"));
    }

    #[test]
    fn report_marks_legacy_project_partial_when_module_stream_is_missing() {
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="PartialProject"
Module=Module1/Module1
Module=MissingMod/MissingMod
"#,
            ),
            ("VBA/Module1", b"Sub AutoOpen()\nEnd Sub\n"),
        ]);
        let config = ParserConfig {
            extract_macro_source: true,
            ..Default::default()
        };
        let app = DocirApp::new(config);
        let parsed = app.parse_bytes(&bytes).expect("parse partial legacy doc");

        let report = app.build_vba_recognition(&parsed, true);
        assert_eq!(report.status, VbaRecognitionStatus::Partial);
        assert_eq!(report.projects.len(), 1);
        assert_eq!(report.projects[0].status, VbaRecognitionStatus::Partial);
        assert_eq!(report.projects[0].modules.len(), 2);

        let extracted = report.projects[0]
            .modules
            .iter()
            .find(|module| module.name == "Module1")
            .expect("extracted module");
        assert_eq!(extracted.status, VbaRecognitionStatus::Extracted);

        let missing = report.projects[0]
            .modules
            .iter()
            .find(|module| module.name == "MissingMod")
            .expect("missing module");
        assert_eq!(missing.status, VbaRecognitionStatus::Partial);
        assert!(missing
            .extraction_errors
            .iter()
            .any(|msg| msg.contains("Missing stream VBA/MissingMod")));
    }

    #[test]
    fn report_extracts_legacy_project_with_mixed_module_kinds() {
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="MixedKinds"
Module=Core/Core
Class=Helper/0
Document=ThisDocument/&H00000000
"#,
            ),
            ("VBA/Core", b"Sub AutoOpen()\nEnd Sub\n"),
            ("VBA/Helper", b"Public Sub Run()\nEnd Sub\n"),
            (
                "VBA/ThisDocument",
                b"Private Sub Document_Open()\nEnd Sub\n",
            ),
        ]);
        let config = ParserConfig {
            extract_macro_source: true,
            ..Default::default()
        };
        let app = DocirApp::new(config);
        let parsed = app.parse_bytes(&bytes).expect("parse mixed legacy doc");

        let report = app.build_vba_recognition(&parsed, true);
        assert_eq!(report.status, VbaRecognitionStatus::Extracted);
        assert_eq!(report.projects.len(), 1);
        assert_eq!(report.projects[0].modules.len(), 3);

        let core = report.projects[0]
            .modules
            .iter()
            .find(|module| module.name == "Core")
            .expect("core module");
        assert_eq!(core.kind, "standard");
        assert_eq!(core.stream_path.as_deref(), Some("VBA/Core"));

        let helper = report.projects[0]
            .modules
            .iter()
            .find(|module| module.name == "Helper")
            .expect("helper class");
        assert_eq!(helper.kind, "class");
        assert_eq!(helper.stream_path.as_deref(), Some("VBA/Helper"));

        let document = report.projects[0]
            .modules
            .iter()
            .find(|module| module.name == "ThisDocument")
            .expect("document module");
        assert_eq!(document.kind, "document");
        assert_eq!(document.stream_path.as_deref(), Some("VBA/ThisDocument"));
        assert!(report.projects[0].has_auto_exec);
    }

    #[test]
    fn report_recognizes_legacy_project_without_requesting_source() {
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="RecognizeOnly"
Module=Core/Core
Class=Helper/0
Document=ThisDocument/&H00000000
"#,
            ),
            ("VBA/Core", b"Sub AutoOpen()\nEnd Sub\n"),
            ("VBA/Helper", b"Public Sub Run()\nEnd Sub\n"),
            (
                "VBA/ThisDocument",
                b"Private Sub Document_Open()\nEnd Sub\n",
            ),
        ]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app
            .parse_bytes(&bytes)
            .expect("parse recognize-only legacy doc");

        let report = app.build_vba_recognition(&parsed, false);
        assert_eq!(report.status, VbaRecognitionStatus::Recognized);
        assert_eq!(report.projects.len(), 1);
        assert_eq!(report.projects[0].status, VbaRecognitionStatus::Recognized);
        assert!(report.projects[0].has_auto_exec);
        assert_eq!(report.projects[0].modules.len(), 3);
        assert!(report.projects[0]
            .modules
            .iter()
            .all(|module| module.status == VbaRecognitionStatus::Recognized));
        assert!(report.projects[0]
            .modules
            .iter()
            .all(|module| module.source_text.is_none()));
    }

    #[test]
    fn report_keeps_partial_status_without_requesting_source() {
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="PartialRecognizeOnly"
Module=Core/Core
Module=MissingMod/MissingMod
"#,
            ),
            ("VBA/Core", b"Sub AutoOpen()\nEnd Sub\n"),
        ]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app
            .parse_bytes(&bytes)
            .expect("parse partial recognize-only doc");

        let report = app.build_vba_recognition(&parsed, false);
        assert_eq!(report.status, VbaRecognitionStatus::Partial);
        assert_eq!(report.projects.len(), 1);
        assert_eq!(report.projects[0].status, VbaRecognitionStatus::Partial);
        assert_eq!(report.projects[0].modules.len(), 2);

        let extracted = report.projects[0]
            .modules
            .iter()
            .find(|module| module.name == "Core")
            .expect("core module");
        assert_eq!(extracted.status, VbaRecognitionStatus::Recognized);
        assert!(extracted.source_text.is_none());

        let missing = report.projects[0]
            .modules
            .iter()
            .find(|module| module.name == "MissingMod")
            .expect("missing module");
        assert_eq!(missing.status, VbaRecognitionStatus::Partial);
        assert!(missing
            .extraction_errors
            .iter()
            .any(|msg| msg.contains("Missing stream VBA/MissingMod")));
    }

    #[test]
    fn report_keeps_error_status_without_requesting_source() {
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="DecodeFailRecognizeOnly"
Module=Broken/Broken
"#,
            ),
            ("VBA/Broken", b""),
        ]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app
            .parse_bytes(&bytes)
            .expect("parse decode-fail recognize-only doc");

        let report = app.build_vba_recognition(&parsed, false);
        assert_eq!(report.status, VbaRecognitionStatus::Error);
        assert_eq!(report.projects.len(), 1);
        assert_eq!(report.projects[0].status, VbaRecognitionStatus::Error);
        assert_eq!(report.projects[0].modules.len(), 1);

        let broken = &report.projects[0].modules[0];
        assert_eq!(broken.name, "Broken");
        assert_eq!(broken.status, VbaRecognitionStatus::Error);
        assert!(broken.source_text.is_none());
        assert!(broken
            .extraction_errors
            .iter()
            .any(|msg| msg.contains("Failed to decompress VBA/Broken")));
    }

    #[test]
    fn report_keeps_protection_and_references_for_partial_project() {
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="ProtectedPartial"
Module=Core/Core
Module=MissingMod/MissingMod
Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation
Reference=*\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#scrrun.dll#Microsoft Scripting Runtime
DPB="AAAA"
"#,
            ),
            ("VBA/Core", b"Sub AutoOpen()\nEnd Sub\n"),
        ]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app
            .parse_bytes(&bytes)
            .expect("parse protected partial doc");

        let report = app.build_vba_recognition(&parsed, false);
        assert_eq!(report.status, VbaRecognitionStatus::Partial);
        assert_eq!(report.projects.len(), 1);
        assert!(report.projects[0].is_protected);
        assert_eq!(report.projects[0].references.len(), 2);
        assert!(report.projects[0].references[0].contains("OLE Automation"));
        assert!(report.projects[0].references[1].contains("Microsoft Scripting Runtime"));
        assert_eq!(report.projects[0].status, VbaRecognitionStatus::Partial);
    }

    #[test]
    fn report_keeps_protection_and_references_for_error_project() {
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="ProtectedError"
Module=Broken/Broken
Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation
DPB="AAAA"
"#,
            ),
            ("VBA/Broken", b""),
        ]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app.parse_bytes(&bytes).expect("parse protected error doc");

        let report = app.build_vba_recognition(&parsed, false);
        assert_eq!(report.status, VbaRecognitionStatus::Error);
        assert_eq!(report.projects.len(), 1);
        assert!(report.projects[0].is_protected);
        assert_eq!(report.projects[0].references.len(), 1);
        assert!(report.projects[0].references[0].contains("OLE Automation"));
        assert_eq!(report.projects[0].status, VbaRecognitionStatus::Error);
    }

    #[test]
    fn report_keeps_mixed_module_kinds_for_partial_protected_project() {
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="MixedProtectedPartial"
Module=Core/Core
Class=Helper/0
Document=ThisDocument/&H00000000
Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation
DPB="AAAA"
"#,
            ),
            ("VBA/Core", b"Sub AutoOpen()\nEnd Sub\n"),
            (
                "VBA/ThisDocument",
                b"Private Sub Document_Open()\nEnd Sub\n",
            ),
        ]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app
            .parse_bytes(&bytes)
            .expect("parse mixed protected partial doc");

        let report = app.build_vba_recognition(&parsed, false);
        assert_eq!(report.status, VbaRecognitionStatus::Partial);
        assert_eq!(report.projects.len(), 1);
        assert!(report.projects[0].is_protected);
        assert_eq!(report.projects[0].references.len(), 1);
        assert_eq!(report.projects[0].modules.len(), 3);

        let core = report.projects[0]
            .modules
            .iter()
            .find(|module| module.name == "Core")
            .expect("core module");
        assert_eq!(core.kind, "standard");
        assert_eq!(core.status, VbaRecognitionStatus::Recognized);

        let helper = report.projects[0]
            .modules
            .iter()
            .find(|module| module.name == "Helper")
            .expect("helper class");
        assert_eq!(helper.kind, "class");
        assert_eq!(helper.status, VbaRecognitionStatus::Partial);
        assert!(helper
            .extraction_errors
            .iter()
            .any(|msg| msg.contains("Missing stream VBA/Helper")));

        let document = report.projects[0]
            .modules
            .iter()
            .find(|module| module.name == "ThisDocument")
            .expect("document module");
        assert_eq!(document.kind, "document");
        assert_eq!(document.status, VbaRecognitionStatus::Recognized);
    }

    #[test]
    fn report_keeps_mixed_module_kinds_for_error_protected_project() {
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="MixedProtectedError"
Module=Core/Core
Document=ThisDocument/&H00000000
Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation
DPB="AAAA"
"#,
            ),
            ("VBA/Core", b"Sub AutoOpen()\nEnd Sub\n"),
            ("VBA/ThisDocument", b""),
        ]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app
            .parse_bytes(&bytes)
            .expect("parse mixed protected error doc");

        let report = app.build_vba_recognition(&parsed, false);
        assert_eq!(report.status, VbaRecognitionStatus::Error);
        assert_eq!(report.projects.len(), 1);
        assert!(report.projects[0].is_protected);
        assert_eq!(report.projects[0].references.len(), 1);
        assert_eq!(report.projects[0].modules.len(), 2);

        let core = report.projects[0]
            .modules
            .iter()
            .find(|module| module.name == "Core")
            .expect("core module");
        assert_eq!(core.kind, "standard");
        assert_eq!(core.status, VbaRecognitionStatus::Recognized);

        let document = report.projects[0]
            .modules
            .iter()
            .find(|module| module.name == "ThisDocument")
            .expect("document module");
        assert_eq!(document.kind, "document");
        assert_eq!(document.status, VbaRecognitionStatus::Error);
        assert!(document
            .extraction_errors
            .iter()
            .any(|msg| msg.contains("Failed to decompress VBA/ThisDocument")));
    }
}
