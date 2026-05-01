use docir_diff::{ChangeKind, DiffEngine};
use docir_parser::DocumentParser;
use std::path::{Path, PathBuf};

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(Path::parent)
        .expect("repository root")
        .to_path_buf()
}

fn parse_fixture(parser: &DocumentParser, relative: &str) -> docir_parser::parser::ParsedDocument {
    let path = repo_root().join(relative);
    parser.parse_file(path).expect("fixture parse")
}

#[test]
fn diff_rich_xlsx_against_itself_is_empty() {
    let parser = DocumentParser::new();
    let left = parse_fixture(&parser, "fixtures/ooxml/rich.xlsx");
    let right = parse_fixture(&parser, "fixtures/ooxml/rich.xlsx");

    let diff = DiffEngine::diff(&left.store, left.root_id, &right.store, right.root_id)
        .expect("diff should succeed");
    assert!(
        diff.is_empty(),
        "diff should be empty for identical documents"
    );
}

#[test]
fn diff_rich_pptx_against_itself_is_empty() {
    let parser = DocumentParser::new();
    let left = parse_fixture(&parser, "fixtures/ooxml/rich.pptx");
    let right = parse_fixture(&parser, "fixtures/ooxml/rich.pptx");

    let diff = DiffEngine::diff(&left.store, left.root_id, &right.store, right.root_id)
        .expect("diff should succeed");
    assert!(
        diff.is_empty(),
        "diff should be empty for identical presentations"
    );
}

#[test]
fn diff_docx_variants_reports_structural_changes() {
    let parser = DocumentParser::new();
    let left = parse_fixture(&parser, "fixtures/ooxml/minimal.docx");
    let right = parse_fixture(&parser, "fixtures/ooxml/rich.docx");

    let diff = DiffEngine::diff(&left.store, left.root_id, &right.store, right.root_id)
        .expect("diff should succeed");
    assert!(
        !diff.added.is_empty() || !diff.removed.is_empty() || !diff.modified.is_empty(),
        "expected differences between minimal and rich docx fixtures"
    );
}

#[test]
fn diff_rtf_variants_include_content_or_metadata_changes() {
    let parser = DocumentParser::new();
    let left = parse_fixture(&parser, "fixtures/rtf/list_styles.rtf");
    let right = parse_fixture(&parser, "fixtures/rtf/object_fields.rtf");

    let diff = DiffEngine::diff(&left.store, left.root_id, &right.store, right.root_id)
        .expect("diff should succeed");
    assert!(
        !diff.is_empty(),
        "expected differences between distinct rtf fixtures"
    );
    assert!(
        diff.modified.iter().any(|m| matches!(
            m.change_kind,
            ChangeKind::Content | ChangeKind::Both | ChangeKind::Metadata
        )) || !diff.added.is_empty()
            || !diff.removed.is_empty(),
        "expected content/metadata modifications or structural add/remove changes"
    );
}

#[test]
fn diff_ooxml_minimal_variants_are_not_empty() {
    let parser = DocumentParser::new();
    let docx = parse_fixture(&parser, "fixtures/ooxml/minimal.docx");
    let xlsx = parse_fixture(&parser, "fixtures/ooxml/minimal.xlsx");
    let pptx = parse_fixture(&parser, "fixtures/ooxml/minimal.pptx");

    let diff_docx_xlsx = DiffEngine::diff(&docx.store, docx.root_id, &xlsx.store, xlsx.root_id)
        .expect("diff should succeed");
    assert!(!diff_docx_xlsx.is_empty());

    let diff_xlsx_pptx = DiffEngine::diff(&xlsx.store, xlsx.root_id, &pptx.store, pptx.root_id)
        .expect("diff should succeed");
    assert!(!diff_xlsx_pptx.is_empty());
}

#[test]
fn diff_rich_odf_variants_are_not_empty() {
    let parser = DocumentParser::new();
    let odt = parse_fixture(&parser, "fixtures/odf/rich.odt");
    let ods = parse_fixture(&parser, "fixtures/odf/minimal.ods");

    let diff = DiffEngine::diff(&odt.store, odt.root_id, &ods.store, ods.root_id)
        .expect("diff should succeed");
    assert!(!diff.is_empty());
}
