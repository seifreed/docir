//! Build a VBA recognition summary without AST or deobfuscation.

use anyhow::Result;
use docir_app::{ExportDocumentRef, ParserConfig, Phase0VbaExport, VbaRecognitionReport};
use std::path::PathBuf;

use crate::cli::JsonOutputOpts;
use crate::commands::util::{
    build_app_and_parse, push_labeled_line, source_format_label, write_json_output,
    write_text_output,
};

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    include_source: bool,
    opts: JsonOutputOpts,
    parser_config: &ParserConfig,
) -> Result<()> {
    let JsonOutputOpts {
        json,
        pretty,
        output,
    } = opts;
    let mut parse_config = parser_config.clone();
    if include_source {
        parse_config.extract_macro_source = true;
    }
    let (app, parsed) = build_app_and_parse(&input, &parse_config)?;
    let report = app.build_vba_recognition(&parsed, include_source);

    if json {
        let export = Phase0VbaExport::from_report(
            &report,
            ExportDocumentRef::new(
                input.display().to_string(),
                source_format_label(&input, parsed.format().extension()),
                Some(input.display().to_string()),
            ),
        );
        return write_json_output(&export, pretty, output);
    }

    let text = format_report_text(&report);
    write_text_output(&text, output)
}

fn format_report_text(report: &VbaRecognitionReport) -> String {
    let mut out = String::new();
    push_labeled_line(&mut out, 0, "Format", &report.document_format);
    push_labeled_line(&mut out, 0, "Status", format!("{:?}", report.status));
    push_labeled_line(&mut out, 0, "Projects", report.projects.len());
    for project in &report.projects {
        out.push_str(&format!(
            "\n- Project {} [{}]\n",
            project.project_name.as_deref().unwrap_or("unnamed"),
            project.node_id
        ));
        if let Some(path) = project.container_path.as_deref() {
            push_labeled_line(&mut out, 2, "Container", path);
        }
        push_labeled_line(&mut out, 2, "Protected", project.is_protected);
        push_labeled_line(&mut out, 2, "AutoExec", project.has_auto_exec);
        if !project.auto_exec_procedures.is_empty() {
            push_labeled_line(
                &mut out,
                2,
                "AutoExec Procedures",
                project.auto_exec_procedures.join(", "),
            );
        }
        push_labeled_line(&mut out, 2, "References", project.references.len());
        for reference in &project.references {
            out.push_str(&format!("    - {}\n", reference));
        }
        push_labeled_line(&mut out, 2, "Modules", project.modules.len());
        for module in &project.modules {
            out.push_str(&format!(
                "    * {} ({}) {:?}\n",
                module.name, module.kind, module.status
            ));
            if let Some(path) = module.stream_path.as_deref() {
                push_labeled_line(&mut out, 6, "Path", path);
            }
            if let Some(encoding) = module.source_encoding.as_deref() {
                push_labeled_line(&mut out, 6, "Encoding", encoding);
            }
            if let Some(sha256) = module.source_hash.as_deref() {
                push_labeled_line(&mut out, 6, "SHA-256", sha256);
            }
            if let Some(compressed) = module.compressed_size {
                push_labeled_line(&mut out, 6, "Compressed", format!("{compressed} bytes"));
            }
            if let Some(decompressed) = module.decompressed_size {
                push_labeled_line(&mut out, 6, "Decompressed", format!("{decompressed} bytes"));
            }
            if !module.procedures.is_empty() {
                push_labeled_line(&mut out, 6, "Procedures", module.procedures.join(", "));
            }
            for error in &module.extraction_errors {
                push_labeled_line(&mut out, 6, "Error", error);
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::format_report_text;
    use crate::commands::util::write_json_output;
    use crate::test_support;
    use docir_app::{
        ExportDocumentRef, Phase0VbaExport, VbaModuleReport, VbaProjectReport,
        VbaRecognitionReport, VbaRecognitionStatus,
    };
    use std::fs;

    #[test]
    fn format_report_text_includes_partial_project_metadata() {
        let report = VbaRecognitionReport {
            document_format: "doc".to_string(),
            status: VbaRecognitionStatus::Partial,
            projects: vec![VbaProjectReport {
                node_id: "node_1".to_string(),
                project_name: Some("MixedProtectedPartial".to_string()),
                container_path: Some("VBA".to_string()),
                storage_root: Some("VBA".to_string()),
                is_protected: true,
                has_auto_exec: true,
                auto_exec_procedures: vec!["AutoOpen".to_string()],
                references: vec!["Reference=...OLE Automation".to_string()],
                status: VbaRecognitionStatus::Partial,
                modules: vec![
                    VbaModuleReport {
                        node_id: "node_2".to_string(),
                        name: "Core".to_string(),
                        kind: "standard".to_string(),
                        stream_name: Some("Core".to_string()),
                        stream_path: Some("VBA/Core".to_string()),
                        status: VbaRecognitionStatus::Recognized,
                        source_text: None,
                        source_encoding: Some("utf-8-lossy-normalized".to_string()),
                        source_hash: Some("a".repeat(64)),
                        compressed_size: Some(24),
                        decompressed_size: Some(24),
                        procedures: vec!["AutoOpen".to_string()],
                        suspicious_calls: Vec::new(),
                        extraction_errors: Vec::new(),
                    },
                    VbaModuleReport {
                        node_id: "node_3".to_string(),
                        name: "Helper".to_string(),
                        kind: "class".to_string(),
                        stream_name: Some("Helper".to_string()),
                        stream_path: Some("VBA/Helper".to_string()),
                        status: VbaRecognitionStatus::Partial,
                        source_text: None,
                        source_encoding: None,
                        source_hash: None,
                        compressed_size: None,
                        decompressed_size: None,
                        procedures: Vec::new(),
                        suspicious_calls: Vec::new(),
                        extraction_errors: vec!["Missing stream VBA/Helper".to_string()],
                    },
                ],
            }],
        };

        let text = format_report_text(&report);
        assert!(text.contains("Format: doc"));
        assert!(text.contains("Status: Partial"));
        assert!(text.contains("- Project MixedProtectedPartial [node_1]"));
        assert!(text.contains("Container: VBA"));
        assert!(text.contains("Protected: true"));
        assert!(text.contains("AutoExec: true"));
        assert!(text.contains("AutoExec Procedures: AutoOpen"));
        assert!(text.contains("References: 1"));
        assert!(text.contains("- Reference=...OLE Automation"));
        assert!(text.contains("* Core (standard) Recognized"));
        assert!(text.contains("Path: VBA/Core"));
        assert!(text.contains("Encoding: utf-8-lossy-normalized"));
        assert!(text.contains(&format!("SHA-256: {}", "a".repeat(64))));
        assert!(text.contains("Compressed: 24 bytes"));
        assert!(text.contains("Decompressed: 24 bytes"));
        assert!(text.contains("Procedures: AutoOpen"));
        assert!(text.contains("* Helper (class) Partial"));
        assert!(text.contains("Path: VBA/Helper"));
        assert!(text.contains("Error: Missing stream VBA/Helper"));
    }

    #[test]
    fn format_report_text_includes_error_details_for_document_module() {
        let report = VbaRecognitionReport {
            document_format: "doc".to_string(),
            status: VbaRecognitionStatus::Error,
            projects: vec![VbaProjectReport {
                node_id: "node_1".to_string(),
                project_name: Some("MixedProtectedError".to_string()),
                container_path: Some("VBA".to_string()),
                storage_root: Some("VBA".to_string()),
                is_protected: true,
                has_auto_exec: true,
                auto_exec_procedures: vec!["AutoOpen".to_string()],
                references: vec!["Reference=...OLE Automation".to_string()],
                status: VbaRecognitionStatus::Error,
                modules: vec![VbaModuleReport {
                    node_id: "node_2".to_string(),
                    name: "ThisDocument".to_string(),
                    kind: "document".to_string(),
                    stream_name: Some("ThisDocument".to_string()),
                    stream_path: Some("VBA/ThisDocument".to_string()),
                    status: VbaRecognitionStatus::Error,
                    source_text: None,
                    source_encoding: None,
                    source_hash: None,
                    compressed_size: Some(0),
                    decompressed_size: None,
                    procedures: Vec::new(),
                    suspicious_calls: Vec::new(),
                    extraction_errors: vec!["Failed to decompress VBA/ThisDocument".to_string()],
                }],
            }],
        };

        let text = format_report_text(&report);
        assert!(text.contains("Status: Error"));
        assert!(text.contains("References: 1"));
        assert!(text.contains("* ThisDocument (document) Error"));
        assert!(text.contains("Path: VBA/ThisDocument"));
        assert!(text.contains("Error: Failed to decompress VBA/ThisDocument"));
    }

    #[test]
    fn write_json_output_for_recognize_vba_export_preserves_cli_visible_fields() {
        let export = Phase0VbaExport::from_report(
            &VbaRecognitionReport {
                document_format: "doc".to_string(),
                status: VbaRecognitionStatus::Error,
                projects: vec![VbaProjectReport {
                    node_id: "node_1".to_string(),
                    project_name: Some("MixedProtectedError".to_string()),
                    container_path: Some("VBA".to_string()),
                    storage_root: Some("VBA".to_string()),
                    is_protected: true,
                    has_auto_exec: true,
                    auto_exec_procedures: vec!["AutoOpen".to_string()],
                    references: vec!["Reference=...OLE Automation".to_string()],
                    status: VbaRecognitionStatus::Error,
                    modules: vec![
                        VbaModuleReport {
                            node_id: "node_2".to_string(),
                            name: "Core".to_string(),
                            kind: "standard".to_string(),
                            stream_name: Some("Core".to_string()),
                            stream_path: Some("VBA/Core".to_string()),
                            status: VbaRecognitionStatus::Recognized,
                            source_text: None,
                            source_encoding: None,
                            source_hash: None,
                            compressed_size: Some(24),
                            decompressed_size: Some(24),
                            procedures: vec!["AutoOpen".to_string()],
                            suspicious_calls: Vec::new(),
                            extraction_errors: Vec::new(),
                        },
                        VbaModuleReport {
                            node_id: "node_3".to_string(),
                            name: "ThisDocument".to_string(),
                            kind: "document".to_string(),
                            stream_name: Some("ThisDocument".to_string()),
                            stream_path: Some("VBA/ThisDocument".to_string()),
                            status: VbaRecognitionStatus::Error,
                            source_text: None,
                            source_encoding: None,
                            source_hash: None,
                            compressed_size: Some(0),
                            decompressed_size: None,
                            procedures: Vec::new(),
                            suspicious_calls: Vec::new(),
                            extraction_errors: vec![
                                "Failed to decompress VBA/ThisDocument".to_string()
                            ],
                        },
                    ],
                }],
            },
            ExportDocumentRef::new("sample.doc", "doc", Some("/tmp/sample.doc".to_string())),
        );
        let output = test_support::temp_file("recognize_vba_export", "json");

        write_json_output(&export, true, Some(output.clone())).expect("write recognize export");

        let json = fs::read_to_string(&output).expect("read recognize export");
        assert!(json.contains("\"status\": \"error\""));
        assert!(json.contains("\"is_protected\": true"));
        assert!(json.contains("\"references\": ["));
        assert!(json.contains("\"kind\": \"document\""));
        assert!(json.contains("\"status\": \"recognized\""));
        assert!(json.contains("\"status\": \"error\""));
        assert!(json.contains("Failed to decompress VBA/ThisDocument"));

        let _ = fs::remove_file(output);
    }
}
