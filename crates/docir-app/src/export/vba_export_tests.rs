use super::*;
use crate::{VbaModuleReport, VbaProjectReport, VbaRecognitionStatus};

#[test]
fn phase0_vba_export_maps_statuses_and_paths() {
    let report = VbaRecognitionReport {
        document_format: "docm".to_string(),
        status: VbaRecognitionStatus::Extracted,
        projects: vec![VbaProjectReport {
            node_id: "node_1".to_string(),
            project_name: Some("Project".to_string()),
            container_path: Some("word/vbaProject.bin".to_string()),
            storage_root: Some("VBA".to_string()),
            is_protected: false,
            has_auto_exec: false,
            auto_exec_procedures: Vec::new(),
            references: Vec::new(),
            status: VbaRecognitionStatus::Extracted,
            modules: vec![VbaModuleReport {
                node_id: "node_2".to_string(),
                name: "Module1".to_string(),
                kind: "standard".to_string(),
                stream_name: Some("Module1".to_string()),
                stream_path: Some("VBA/Module1".to_string()),
                status: VbaRecognitionStatus::Extracted,
                source_text: Some("Sub Test()\nEnd Sub\n".to_string()),
                source_encoding: Some("utf-8".to_string()),
                source_hash: Some("a".repeat(64)),
                compressed_size: Some(10),
                decompressed_size: Some(18),
                procedures: Vec::new(),
                suspicious_calls: Vec::new(),
                extraction_errors: Vec::new(),
            }],
        }],
    };

    let export = Phase0VbaExport::from_report(
        &report,
        ExportDocumentRef::new("doc", "docm", Some("/tmp/doc.docm".to_string())),
    );
    assert_eq!(export.vba.status, "extracted");
    assert_eq!(
        export.vba.projects[0].modules[0].storage_path,
        "VBA/Module1"
    );
    assert!(!export.vba.projects[0].is_protected);
}

#[test]
fn phase0_vba_export_omits_source_text_for_recognition_only() {
    let report = VbaRecognitionReport {
        document_format: "doc".to_string(),
        status: VbaRecognitionStatus::Recognized,
        projects: vec![VbaProjectReport {
            node_id: "node_1".to_string(),
            project_name: Some("RecognizeOnly".to_string()),
            container_path: Some("VBA".to_string()),
            storage_root: Some("VBA".to_string()),
            is_protected: false,
            has_auto_exec: true,
            auto_exec_procedures: vec!["AutoOpen".to_string()],
            references: Vec::new(),
            status: VbaRecognitionStatus::Recognized,
            modules: vec![VbaModuleReport {
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
            }],
        }],
    };

    let export = Phase0VbaExport::from_report(
        &report,
        ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
    );

    assert_eq!(export.vba.status, "recognized");
    assert_eq!(export.vba.projects[0].modules[0].status, "recognized");
    assert!(export.vba.projects[0].modules[0].source_text.is_none());
    assert!(export.vba.projects[0].modules[0].source_encoding.is_none());
    assert!(!export.vba.projects[0].is_protected);
}

#[test]
fn phase0_vba_export_keeps_source_text_when_extracted() {
    let report = VbaRecognitionReport {
        document_format: "doc".to_string(),
        status: VbaRecognitionStatus::Extracted,
        projects: vec![VbaProjectReport {
            node_id: "node_1".to_string(),
            project_name: Some("RecognizeWithSource".to_string()),
            container_path: Some("VBA".to_string()),
            storage_root: Some("VBA".to_string()),
            is_protected: false,
            has_auto_exec: true,
            auto_exec_procedures: vec!["AutoOpen".to_string()],
            references: Vec::new(),
            status: VbaRecognitionStatus::Extracted,
            modules: vec![VbaModuleReport {
                node_id: "node_2".to_string(),
                name: "Core".to_string(),
                kind: "standard".to_string(),
                stream_name: Some("Core".to_string()),
                stream_path: Some("VBA/Core".to_string()),
                status: VbaRecognitionStatus::Extracted,
                source_text: Some("Sub AutoOpen()\nEnd Sub\n".to_string()),
                source_encoding: Some("utf-8-lossy-normalized".to_string()),
                source_hash: Some("deadbeef".to_string()),
                compressed_size: Some(24),
                decompressed_size: Some(24),
                procedures: vec!["AutoOpen".to_string()],
                suspicious_calls: Vec::new(),
                extraction_errors: Vec::new(),
            }],
        }],
    };

    let export = Phase0VbaExport::from_report(
        &report,
        ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
    );

    assert_eq!(export.vba.status, "extracted");
    assert_eq!(export.vba.projects[0].modules[0].status, "extracted");
    assert_eq!(
        export.vba.projects[0].modules[0].source_text.as_deref(),
        Some("Sub AutoOpen()\nEnd Sub\n")
    );
    assert_eq!(
        export.vba.projects[0].modules[0].source_encoding.as_deref(),
        Some("compressed-container-decoded")
    );
    assert!(!export.vba.projects[0].is_protected);
}

#[test]
fn phase0_vba_export_keeps_partial_status_without_source_text() {
    let report = VbaRecognitionReport {
        document_format: "doc".to_string(),
        status: VbaRecognitionStatus::Partial,
        projects: vec![VbaProjectReport {
            node_id: "node_1".to_string(),
            project_name: Some("PartialRecognizeOnly".to_string()),
            container_path: Some("VBA".to_string()),
            storage_root: Some("VBA".to_string()),
            is_protected: false,
            has_auto_exec: true,
            auto_exec_procedures: vec!["AutoOpen".to_string()],
            references: Vec::new(),
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
                    name: "MissingMod".to_string(),
                    kind: "standard".to_string(),
                    stream_name: Some("MissingMod".to_string()),
                    stream_path: Some("VBA/MissingMod".to_string()),
                    status: VbaRecognitionStatus::Partial,
                    source_text: None,
                    source_encoding: None,
                    source_hash: None,
                    compressed_size: None,
                    decompressed_size: None,
                    procedures: Vec::new(),
                    suspicious_calls: Vec::new(),
                    extraction_errors: vec!["Missing stream VBA/MissingMod".to_string()],
                },
            ],
        }],
    };

    let export = Phase0VbaExport::from_report(
        &report,
        ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
    );

    assert_eq!(export.vba.status, "partial");
    assert_eq!(export.vba.projects[0].status, "partial");
    assert_eq!(export.vba.projects[0].modules[0].status, "recognized");
    assert_eq!(export.vba.projects[0].modules[1].status, "partial");
    assert!(export.vba.projects[0].modules[1].source_text.is_none());
    assert_eq!(export.vba.projects[0].modules[1].diagnostics.len(), 1);
    assert!(!export.vba.projects[0].is_protected);
}

#[test]
fn phase0_vba_export_keeps_error_status_without_source_text() {
    let report = VbaRecognitionReport {
        document_format: "doc".to_string(),
        status: VbaRecognitionStatus::Error,
        projects: vec![VbaProjectReport {
            node_id: "node_1".to_string(),
            project_name: Some("DecodeFailRecognizeOnly".to_string()),
            container_path: Some("VBA".to_string()),
            storage_root: Some("VBA".to_string()),
            is_protected: false,
            has_auto_exec: false,
            auto_exec_procedures: Vec::new(),
            references: Vec::new(),
            status: VbaRecognitionStatus::Error,
            modules: vec![VbaModuleReport {
                node_id: "node_2".to_string(),
                name: "Broken".to_string(),
                kind: "standard".to_string(),
                stream_name: Some("Broken".to_string()),
                stream_path: Some("VBA/Broken".to_string()),
                status: VbaRecognitionStatus::Error,
                source_text: None,
                source_encoding: None,
                source_hash: None,
                compressed_size: Some(0),
                decompressed_size: None,
                procedures: Vec::new(),
                suspicious_calls: Vec::new(),
                extraction_errors: vec!["Failed to decompress VBA/Broken".to_string()],
            }],
        }],
    };

    let export = Phase0VbaExport::from_report(
        &report,
        ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
    );

    assert_eq!(export.vba.status, "error");
    assert_eq!(export.vba.projects[0].status, "error");
    assert_eq!(export.vba.projects[0].modules[0].status, "error");
    assert!(export.vba.projects[0].modules[0].source_text.is_none());
    assert_eq!(export.vba.projects[0].modules[0].diagnostics.len(), 1);
    assert_eq!(
        export.vba.projects[0].modules[0].diagnostics[0].message,
        "Failed to decompress VBA/Broken"
    );
    assert!(!export.vba.projects[0].is_protected);
}

#[test]
fn phase0_vba_export_preserves_protection_and_references_for_partial_project() {
    let report = VbaRecognitionReport {
        document_format: "doc".to_string(),
        status: VbaRecognitionStatus::Partial,
        projects: vec![VbaProjectReport {
            node_id: "node_1".to_string(),
            project_name: Some("ProtectedPartial".to_string()),
            container_path: Some("VBA".to_string()),
            storage_root: Some("VBA".to_string()),
            is_protected: true,
            has_auto_exec: true,
            auto_exec_procedures: vec!["AutoOpen".to_string()],
            references: vec![
                r#"Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation"#.to_string(),
                r#"Reference=*\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#scrrun.dll#Microsoft Scripting Runtime"#.to_string(),
            ],
            status: VbaRecognitionStatus::Partial,
            modules: vec![VbaModuleReport {
                node_id: "node_2".to_string(),
                name: "Core".to_string(),
                kind: "standard".to_string(),
                stream_name: Some("Core".to_string()),
                stream_path: Some("VBA/Core".to_string()),
                status: VbaRecognitionStatus::Partial,
                source_text: None,
                source_encoding: None,
                source_hash: None,
                compressed_size: Some(24),
                decompressed_size: None,
                procedures: vec!["AutoOpen".to_string()],
                suspicious_calls: Vec::new(),
                extraction_errors: vec!["Missing stream VBA/Core".to_string()],
            }],
        }],
    };

    let export = Phase0VbaExport::from_report(
        &report,
        ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
    );

    assert!(export.vba.projects[0].is_protected);
    assert_eq!(export.vba.projects[0].references.len(), 2);
    assert!(export.vba.projects[0].references[0].contains("OLE Automation"));
    assert!(export.vba.projects[0].references[1].contains("Microsoft Scripting Runtime"));
}

#[test]
fn phase0_vba_export_preserves_mixed_module_kinds_for_partial_project() {
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
            references: vec![
                r#"Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation"#.to_string(),
            ],
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
                VbaModuleReport {
                    node_id: "node_4".to_string(),
                    name: "ThisDocument".to_string(),
                    kind: "document".to_string(),
                    stream_name: Some("ThisDocument".to_string()),
                    stream_path: Some("VBA/ThisDocument".to_string()),
                    status: VbaRecognitionStatus::Recognized,
                    source_text: None,
                    source_encoding: None,
                    source_hash: None,
                    compressed_size: Some(32),
                    decompressed_size: Some(32),
                    procedures: vec!["Document_Open".to_string()],
                    suspicious_calls: Vec::new(),
                    extraction_errors: Vec::new(),
                },
            ],
        }],
    };

    let export = Phase0VbaExport::from_report(
        &report,
        ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
    );

    assert_eq!(export.vba.status, "partial");
    assert!(export.vba.projects[0].is_protected);
    assert_eq!(export.vba.projects[0].references.len(), 1);
    assert_eq!(export.vba.projects[0].modules[0].kind, "standard");
    assert_eq!(export.vba.projects[0].modules[1].kind, "class");
    assert_eq!(export.vba.projects[0].modules[1].status, "partial");
    assert_eq!(export.vba.projects[0].modules[2].kind, "document");
    assert_eq!(export.vba.projects[0].modules[2].status, "recognized");
}

#[test]
fn phase0_vba_export_preserves_mixed_module_kinds_for_error_project() {
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
            references: vec![
                r#"Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation"#.to_string(),
            ],
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
                    extraction_errors: vec!["Failed to decompress VBA/ThisDocument".to_string()],
                },
            ],
        }],
    };

    let export = Phase0VbaExport::from_report(
        &report,
        ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
    );

    assert_eq!(export.vba.status, "error");
    assert!(export.vba.projects[0].is_protected);
    assert_eq!(export.vba.projects[0].references.len(), 1);
    assert_eq!(export.vba.projects[0].modules[0].kind, "standard");
    assert_eq!(export.vba.projects[0].modules[0].status, "recognized");
    assert_eq!(export.vba.projects[0].modules[1].kind, "document");
    assert_eq!(export.vba.projects[0].modules[1].status, "error");
    assert_eq!(export.vba.projects[0].modules[1].diagnostics.len(), 1);
}
