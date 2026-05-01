use docir_core::ir::IRNode;
use docir_core::types::NodeType;
use docir_parser::{OdfParser, ParseError, ParserConfig};
use std::io::{Cursor, Write};

fn build_zip(entries: &[(&str, &[u8])]) -> Result<Vec<u8>, ParseError> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::SimpleFileOptions::default();
    for (path, data) in entries {
        zip.start_file(path, options)
            .map_err(|e| ParseError::InvalidZip(format!("zip start_file failed: {e}")))?;
        zip.write_all(data)
            .map_err(|e| ParseError::InvalidZip(format!("zip write_all failed: {e}")))?;
    }
    zip.finish()
        .map_err(|e| ParseError::InvalidZip(format!("zip finish failed: {e}")))
        .map(|c| c.into_inner())
}

#[test]
fn parse_odt_inline_events_populate_bookmarks_fields_and_paragraphs() {
    let content_xml = br#"
        <office:document-content xmlns:office="office" xmlns:text="text" xmlns:table="table" xmlns:draw="draw">
          <office:body>
            <office:text>
              <text:list text:style-name="L1">
                <text:list-item>
                  <text:p text:outline-level="2">Hello<text:s text:c="2"/><text:tab/><text:line-break/><text:bookmark-start text:name="bm1"/><text:bookmark-end text:name="bm1"/><text:date/><text:time/></text:p>
                </text:list-item>
              </text:list>
              <text:h text:outline-level="1">Heading</text:h>
            </office:text>
          </office:body>
        </office:document-content>
    "#;
    let zip_bytes = build_zip(&[
        ("mimetype", b"application/vnd.oasis.opendocument.text"),
        ("content.xml", content_xml),
    ])
    .expect("zip");

    let parser = OdfParser::new();
    let parsed = parser
        .parse_reader(Cursor::new(zip_bytes))
        .expect("odt parse");
    let doc = parsed.document().expect("document");
    assert_eq!(doc.format.display_name(), "OpenDocument Text");
    assert!(!doc.content.is_empty());

    let paragraph_count = parsed.store.iter_ids_by_type(NodeType::Paragraph).count();
    let bookmark_start_count = parsed
        .store
        .iter_ids_by_type(NodeType::BookmarkStart)
        .count();
    let bookmark_end_count = parsed.store.iter_ids_by_type(NodeType::BookmarkEnd).count();
    let field_count = parsed.store.iter_ids_by_type(NodeType::Field).count();

    assert!(paragraph_count >= 2, "expected paragraph and heading nodes");
    assert!(bookmark_start_count >= 1);
    assert!(bookmark_end_count >= 1);
    assert!(field_count >= 2, "expected date/time fields");
}

#[test]
fn parse_ods_fast_sampling_handles_sample_limits_and_repeated_cells() {
    let content_xml = br#"
        <office:document-content xmlns:office="office" xmlns:table="table" xmlns:text="text">
          <office:body>
            <office:spreadsheet>
              <table:table table:name="Sheet1">
                <table:table-row>
                  <table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell>
                  <table:table-cell table:number-columns-repeated="5" office:value-type="float" office:value="1"/>
                  <table:covered-table-cell table:number-columns-repeated="2"/>
                </table:table-row>
              </table:table>
            </office:spreadsheet>
          </office:body>
        </office:document-content>
    "#;
    let zip_bytes = build_zip(&[
        (
            "mimetype",
            b"application/vnd.oasis.opendocument.spreadsheet",
        ),
        ("content.xml", content_xml),
    ])
    .expect("zip");

    let mut config = ParserConfig::default();
    config.odf.force_fast = true;
    config.odf.fast_sample_rows = 1;
    config.odf.fast_sample_cols = 1;

    let parser = OdfParser::with_config(config);
    let parsed = parser
        .parse_reader(Cursor::new(zip_bytes))
        .expect("ods parse");
    let doc = parsed.document().expect("document");
    assert_eq!(doc.format.display_name(), "OpenDocument Spreadsheet");
    assert!(!doc.content.is_empty());

    let worksheet_count = parsed.store.iter_ids_by_type(NodeType::Worksheet).count();
    assert!(worksheet_count >= 1);
}

#[test]
fn parse_ods_fast_sampling_reports_xml_error_for_malformed_content() {
    let malformed = br#"
        <office:document-content xmlns:office="office" xmlns:table="table" xmlns:text="text">
          <office:body>
            <office:spreadsheet>
              <table:table table:name="Broken">
                <table:table-row>
                  <table:table-cell office:value-type="string"><text:p>oops
              </table:table>
            </office:spreadsheet>
          </office:body>
        </office:document-content>
    "#;
    let zip_bytes = build_zip(&[
        (
            "mimetype",
            b"application/vnd.oasis.opendocument.spreadsheet",
        ),
        ("content.xml", malformed),
    ])
    .expect("zip");

    let mut config = ParserConfig::default();
    config.odf.force_fast = true;
    config.odf.fast_sample_rows = 1;
    config.odf.fast_sample_cols = 1;
    let parser = OdfParser::with_config(config);

    let err = parser
        .parse_reader(Cursor::new(zip_bytes))
        .expect_err("malformed xml should fail");
    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn parse_odt_with_unsupported_mimetype_fails() {
    let zip_bytes = build_zip(&[
        ("mimetype", b"application/not-odf"),
        ("content.xml", b"<office:document-content/>"),
    ])
    .expect("zip");
    let parser = OdfParser::new();
    let err = parser
        .parse_reader(Cursor::new(zip_bytes))
        .expect_err("unsupported mimetype must fail");
    assert!(matches!(err, ParseError::UnsupportedFormat(_)));
}

#[test]
fn parse_odt_without_mimetype_fails() {
    let zip_bytes = build_zip(&[("content.xml", b"<office:document-content/>")]).expect("zip");
    let parser = OdfParser::new();
    let err = parser
        .parse_reader(Cursor::new(zip_bytes))
        .expect_err("missing mimetype must fail");
    assert!(matches!(err, ParseError::UnsupportedFormat(_)));
}

#[test]
fn parse_ods_fast_mode_emits_section_and_cells() {
    let content_xml = br#"
        <office:document-content xmlns:office="office" xmlns:table="table" xmlns:text="text">
          <office:body>
            <office:spreadsheet>
              <table:table table:name="SheetZ">
                <table:table-row>
                  <table:table-cell office:value-type="string"><text:p>V</text:p></table:table-cell>
                </table:table-row>
              </table:table>
            </office:spreadsheet>
          </office:body>
        </office:document-content>
    "#;
    let zip_bytes = build_zip(&[
        (
            "mimetype",
            b"application/vnd.oasis.opendocument.spreadsheet",
        ),
        ("content.xml", content_xml),
    ])
    .expect("zip");
    let mut config = ParserConfig::default();
    config.odf.force_fast = true;
    config.odf.fast_sample_rows = 1;
    config.odf.fast_sample_cols = 2;

    let parser = OdfParser::with_config(config);
    let parsed = parser
        .parse_reader(Cursor::new(zip_bytes))
        .expect("ods parse");
    let cell_count = parsed.store.iter_ids_by_type(NodeType::Cell).count();
    assert!(cell_count >= 1);

    let has_sheet = parsed.store.values().any(|node| match node {
        IRNode::Worksheet(sheet) => sheet.name == "SheetZ",
        _ => false,
    });
    assert!(has_sheet);
}

#[test]
fn parse_ods_fast_sampling_exercises_start_cell_skip_paths() {
    let content_xml = br#"
        <office:document-content xmlns:office="office" xmlns:table="table" xmlns:text="text">
          <office:body>
            <office:spreadsheet>
              <table:table table:name="SkipPaths">
                <table:table-row>
                  <table:table-cell office:value-type="string"><text:p>keep</text:p></table:table-cell>
                  <table:table-cell table:number-columns-repeated="3"><text:p>skip-me</text:p></table:table-cell>
                  <table:covered-table-cell table:number-columns-repeated="2"><text:p>covered-skip</text:p></table:covered-table-cell>
                </table:table-row>
              </table:table>
            </office:spreadsheet>
          </office:body>
        </office:document-content>
    "#;
    let zip_bytes = build_zip(&[
        (
            "mimetype",
            b"application/vnd.oasis.opendocument.spreadsheet",
        ),
        ("content.xml", content_xml),
    ])
    .expect("zip");

    let mut config = ParserConfig::default();
    config.odf.force_fast = true;
    config.odf.fast_sample_rows = 1;
    config.odf.fast_sample_cols = 1;

    let parser = OdfParser::with_config(config);
    let parsed = parser
        .parse_reader(Cursor::new(zip_bytes))
        .expect("ods parse");
    let worksheet_id = parsed
        .store
        .iter_ids_by_type(NodeType::Worksheet)
        .next()
        .expect("worksheet id");
    let worksheet = parsed
        .store
        .get(worksheet_id)
        .and_then(|node| match node {
            IRNode::Worksheet(w) => Some(w),
            _ => None,
        })
        .expect("worksheet");
    assert!(
        worksheet.cells.len() <= 1,
        "fast sampling should keep at most one sampled cell"
    );
}

#[test]
fn parse_ods_fast_sampling_exercises_covered_cell_parse_paths_under_limit() {
    let content_xml = br#"
        <office:document-content xmlns:office="office" xmlns:table="table" xmlns:text="text">
          <office:body>
            <office:spreadsheet>
              <table:table table:name="CoveredPaths">
                <table:table-row>
                  <table:covered-table-cell table:number-columns-repeated="2"><text:p>x</text:p></table:covered-table-cell>
                  <table:table-cell table:number-columns-repeated="2"><text:p>value</text:p></table:table-cell>
                  <table:covered-table-cell table:number-columns-repeated="2"/>
                </table:table-row>
              </table:table>
            </office:spreadsheet>
          </office:body>
        </office:document-content>
    "#;
    let zip_bytes = build_zip(&[
        (
            "mimetype",
            b"application/vnd.oasis.opendocument.spreadsheet",
        ),
        ("content.xml", content_xml),
    ])
    .expect("zip");

    let mut config = ParserConfig::default();
    config.odf.force_fast = true;
    config.odf.fast_sample_rows = 1;
    config.odf.fast_sample_cols = 3;

    let parser = OdfParser::with_config(config);
    let parsed = parser
        .parse_reader(Cursor::new(zip_bytes))
        .expect("ods parse");
    let cell_count = parsed.store.iter_ids_by_type(NodeType::Cell).count();
    assert!(cell_count >= 1, "expected sampled cells to be materialized");
}

#[test]
fn parse_odt_enforces_paragraph_limit() {
    let content_xml = br#"
        <office:document-content xmlns:office="office" xmlns:text="text">
          <office:body>
            <office:text>
              <text:p>one</text:p>
              <text:p>two</text:p>
            </office:text>
          </office:body>
        </office:document-content>
    "#;
    let zip_bytes = build_zip(&[
        ("mimetype", b"application/vnd.oasis.opendocument.text"),
        ("content.xml", content_xml),
    ])
    .expect("zip");

    let mut config = ParserConfig::default();
    config.odf.max_paragraphs = Some(1);
    let parser = OdfParser::with_config(config);
    let err = parser
        .parse_reader(Cursor::new(zip_bytes))
        .expect_err("paragraph limit should fail");
    match err {
        ParseError::ResourceLimit(msg) => assert!(msg.contains("paragraphs")),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn parse_ods_enforces_row_limit() {
    let content_xml = br#"
        <office:document-content xmlns:office="office" xmlns:table="table" xmlns:text="text">
          <office:body>
            <office:spreadsheet>
              <table:table table:name="Rows">
                <table:table-row><table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell></table:table-row>
                <table:table-row><table:table-cell office:value-type="string"><text:p>B</text:p></table:table-cell></table:table-row>
              </table:table>
            </office:spreadsheet>
          </office:body>
        </office:document-content>
    "#;
    let zip_bytes = build_zip(&[
        (
            "mimetype",
            b"application/vnd.oasis.opendocument.spreadsheet",
        ),
        ("content.xml", content_xml),
    ])
    .expect("zip");

    let mut config = ParserConfig::default();
    config.odf.max_rows = Some(1);
    let parser = OdfParser::with_config(config);
    let err = parser
        .parse_reader(Cursor::new(zip_bytes))
        .expect_err("row limit should fail");
    match err {
        ParseError::ResourceLimit(msg) => assert!(msg.contains("rows")),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn parse_ods_enforces_content_xml_size_limit() {
    let content_xml = br#"
        <office:document-content xmlns:office="office" xmlns:table="table" xmlns:text="text">
          <office:body><office:spreadsheet><table:table table:name="Big"/></office:spreadsheet></office:body>
        </office:document-content>
    "#;
    let zip_bytes = build_zip(&[
        (
            "mimetype",
            b"application/vnd.oasis.opendocument.spreadsheet",
        ),
        ("content.xml", content_xml),
    ])
    .expect("zip");

    let mut config = ParserConfig::default();
    config.odf.max_bytes = Some(8);
    let parser = OdfParser::with_config(config);
    let err = parser
        .parse_reader(Cursor::new(zip_bytes))
        .expect_err("content size limit should fail");
    match err {
        ParseError::ResourceLimit(msg) => assert!(msg.contains("content.xml too large")),
        other => panic!("unexpected error: {other:?}"),
    }
}
