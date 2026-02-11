use std::fs;
use std::path::PathBuf;

use docir_core::ir::{DiagnosticSeverity, IRNode};
use docir_core::types::DocumentFormat;
use docir_core::types::NodeType;
use docir_parser::DocumentParser;

#[derive(serde::Deserialize)]
struct Manifest {
    fixtures: Vec<FixtureEntry>,
}

#[derive(serde::Deserialize)]
struct FixtureEntry {
    path: String,
    format: String,
}

fn parse_format(value: &str) -> DocumentFormat {
    match value.to_ascii_uppercase().as_str() {
        "DOCX" => DocumentFormat::WordProcessing,
        "XLSX" => DocumentFormat::Spreadsheet,
        "PPTX" => DocumentFormat::Presentation,
        "ODT" => DocumentFormat::OdfText,
        "ODS" => DocumentFormat::OdfSpreadsheet,
        "ODP" => DocumentFormat::OdfPresentation,
        "HWP" => DocumentFormat::Hwp,
        "HWPX" => DocumentFormat::Hwpx,
        "RTF" => DocumentFormat::Rtf,
        other => panic!("Unknown format in manifest: {other}"),
    }
}

#[test]
fn parse_all_fixtures() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .to_path_buf();
    let manifest_path = root.join("fixtures/manifest.json");
    let data = fs::read_to_string(&manifest_path).expect("read manifest");
    let manifest: Manifest = serde_json::from_str(&data).expect("parse manifest");

    let parser = DocumentParser::new();
    for fixture in manifest.fixtures {
        let path = root.join(&fixture.path);
        let expected = parse_format(&fixture.format);
        let parsed = parser.parse_file(&path).expect("parse fixture");
        assert_eq!(parsed.format, expected, "format mismatch for {:?}", path);
        assert!(
            parsed.document().is_some(),
            "missing document node for {:?}",
            path
        );

        // Rich fixtures: assert presence of key nodes (semantic coverage).
        let file_name = path.file_name().and_then(|s| s.to_str()).unwrap_or("");
        if file_name.starts_with("rich.") || file_name == "rich.rtf" || file_name == "rich.hwpx" {
            match expected {
                DocumentFormat::WordProcessing => {
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::StyleSet)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::NumberingSet)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::Comment)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::Footnote)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::Header)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::Footer)
                        .next()
                        .is_some());
                }
                DocumentFormat::Spreadsheet => {
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::Worksheet)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::SharedStringTable)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::SpreadsheetStyles)
                        .next()
                        .is_some());
                }
                DocumentFormat::Presentation => {
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::Slide)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::SlideMaster)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::SlideLayout)
                        .next()
                        .is_some());
                }
                DocumentFormat::OdfText => {
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::StyleSet)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::Comment)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::Footnote)
                        .next()
                        .is_some());
                }
                DocumentFormat::Hwpx => {
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::StyleSet)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::Header)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::Footer)
                        .next()
                        .is_some());
                }
                DocumentFormat::Rtf => {
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::StyleSet)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::Table)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::Hyperlink)
                        .next()
                        .is_some());
                    assert!(parsed
                        .store
                        .iter_ids_by_type(NodeType::OleObject)
                        .next()
                        .is_some());
                }
                _ => {}
            }
        }

        // For OOXML packages, ensure there are no pending matched parts in coverage diagnostics.
        if (file_name.starts_with("rich.")
            || file_name == "rich.docx"
            || file_name == "rich.xlsx"
            || file_name == "rich.pptx")
            && matches!(
                expected,
                DocumentFormat::WordProcessing
                    | DocumentFormat::Spreadsheet
                    | DocumentFormat::Presentation
            )
        {
            if let Some(doc) = parsed.document() {
                let mut pending_parts = 0usize;
                for diag_id in &doc.diagnostics {
                    if let Some(IRNode::Diagnostics(diag)) = parsed.store.get(*diag_id) {
                        for entry in &diag.entries {
                            if entry.code == "COVERAGE_PART"
                                && matches!(entry.severity, DiagnosticSeverity::Warning)
                            {
                                pending_parts += 1;
                            }
                        }
                    }
                }
                assert_eq!(
                    pending_parts, 0,
                    "OOXML fixture has pending parts (unparsed content) for {:?}",
                    path
                );
            }
        }
    }
}
