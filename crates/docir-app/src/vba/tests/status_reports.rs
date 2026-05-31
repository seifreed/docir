use super::*;
use crate::test_support::build_test_cfb;
use crate::{DocirApp, ParserConfig};

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
    assert!(
        report.projects[0]
            .modules
            .iter()
            .all(|module| module.status == VbaRecognitionStatus::Recognized)
    );
    assert!(
        report.projects[0]
            .modules
            .iter()
            .all(|module| module.source_text.is_none())
    );
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
    assert!(
        missing
            .extraction_errors
            .iter()
            .any(|msg| msg.contains("Missing stream VBA/MissingMod"))
    );
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
    assert!(
        broken
            .extraction_errors
            .iter()
            .any(|msg| msg.contains("Failed to decompress VBA/Broken"))
    );
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
    assert!(
        helper
            .extraction_errors
            .iter()
            .any(|msg| msg.contains("Missing stream VBA/Helper"))
    );

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
    assert!(
        document
            .extraction_errors
            .iter()
            .any(|msg| msg.contains("Failed to decompress VBA/ThisDocument"))
    );
}
