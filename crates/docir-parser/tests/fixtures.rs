use std::fs;
use std::path::PathBuf;

use docir_core::ir::{DiagnosticEntry, DiagnosticSeverity, IRNode};
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
        "XLSB" => DocumentFormat::Spreadsheet,
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

fn has_node_type(parsed: &docir_parser::parser::ParsedDocument, node_type: NodeType) -> bool {
    parsed.store.iter_ids_by_type(node_type).next().is_some()
}

fn assert_rich_fixture_nodes(
    parsed: &docir_parser::parser::ParsedDocument,
    expected: DocumentFormat,
) {
    match expected {
        DocumentFormat::WordProcessing => {
            assert!(has_node_type(parsed, NodeType::StyleSet));
            assert!(has_node_type(parsed, NodeType::NumberingSet));
            assert!(has_node_type(parsed, NodeType::Comment));
            assert!(has_node_type(parsed, NodeType::Footnote));
            assert!(has_node_type(parsed, NodeType::Header));
            assert!(has_node_type(parsed, NodeType::Footer));
        }
        DocumentFormat::Spreadsheet => {
            assert!(has_node_type(parsed, NodeType::Worksheet));
            assert!(has_node_type(parsed, NodeType::SharedStringTable));
            assert!(has_node_type(parsed, NodeType::SpreadsheetStyles));
        }
        DocumentFormat::Presentation => {
            assert!(has_node_type(parsed, NodeType::Slide));
            assert!(has_node_type(parsed, NodeType::SlideMaster));
            assert!(has_node_type(parsed, NodeType::SlideLayout));
        }
        DocumentFormat::OdfText => {
            assert!(has_node_type(parsed, NodeType::StyleSet));
            assert!(has_node_type(parsed, NodeType::Comment));
            assert!(has_node_type(parsed, NodeType::Footnote));
        }
        DocumentFormat::Hwpx => {
            assert!(has_node_type(parsed, NodeType::StyleSet));
            assert!(has_node_type(parsed, NodeType::Header));
            assert!(has_node_type(parsed, NodeType::Footer));
        }
        DocumentFormat::Rtf => {
            assert!(has_node_type(parsed, NodeType::StyleSet));
            assert!(has_node_type(parsed, NodeType::Table));
            assert!(has_node_type(parsed, NodeType::Hyperlink));
            assert!(has_node_type(parsed, NodeType::OleObject));
        }
        _ => {}
    }
}

fn assert_no_pending_ooxml_parts(parsed: &docir_parser::parser::ParsedDocument, path: &PathBuf) {
    let Some(doc) = parsed.document() else {
        return;
    };

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

fn collect_diagnostic_entries(
    parsed: &docir_parser::parser::ParsedDocument,
) -> Vec<DiagnosticEntry> {
    let Some(doc) = parsed.document() else {
        return Vec::new();
    };

    let mut entries = Vec::new();
    for diag_id in &doc.diagnostics {
        if let Some(IRNode::Diagnostics(diag)) = parsed.store.get(*diag_id) {
            entries.extend(diag.entries.clone());
        }
    }
    entries
}

fn assert_core_fixture_behavior(
    parsed: &docir_parser::parser::ParsedDocument,
    expected: DocumentFormat,
    path: &PathBuf,
) {
    let doc = parsed
        .document()
        .expect("document node should exist for parsed fixture");
    assert!(
        !doc.content.is_empty(),
        "fixture should produce at least one top-level content node: {:?}",
        path
    );

    let diagnostics = collect_diagnostic_entries(parsed);
    if matches!(
        expected,
        DocumentFormat::WordProcessing | DocumentFormat::Spreadsheet | DocumentFormat::Presentation
    ) {
        assert!(
            diagnostics.iter().any(|entry| entry.code == "COVERAGE_SUMMARY"),
            "OOXML fixture should emit COVERAGE_SUMMARY diagnostics: {:?}",
            path
        );
        assert!(
            diagnostics.iter().any(|entry| entry.code == "COVERAGE_COUNTS"),
            "OOXML fixture should emit COVERAGE_COUNTS diagnostics: {:?}",
            path
        );
        assert!(
            diagnostics.iter().any(|entry| {
                entry.code == "COVERAGE_PART"
                    && matches!(entry.severity, DiagnosticSeverity::Info | DiagnosticSeverity::Warning)
            }),
            "OOXML fixture should emit COVERAGE_PART diagnostics: {:?}",
            path
        );
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
        assert_core_fixture_behavior(&parsed, expected, &path);

        // Rich fixtures: assert presence of key nodes (semantic coverage).
        let file_name = path.file_name().and_then(|s| s.to_str()).unwrap_or("");
        if file_name.starts_with("rich.") || file_name == "rich.rtf" || file_name == "rich.hwpx" {
            assert_rich_fixture_nodes(&parsed, expected);
        }

        if file_name == "object_fields.rtf" {
            assert!(
                has_node_type(&parsed, NodeType::OleObject),
                "object_fields.rtf should surface OLE object nodes"
            );
            assert!(
                has_node_type(&parsed, NodeType::Hyperlink),
                "object_fields.rtf should surface hyperlink field nodes"
            );
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
            assert_no_pending_ooxml_parts(&parsed, &path);
        }
    }
}
