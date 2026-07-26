use super::tests_prelude::*;
use crate::error::ParseError;
use crate::parser::DocumentParser;
use docir_core::security::ThreatIndicatorType;
use std::io::{Cursor, Write};
use zip::ZipWriter;
use zip::write::FileOptions;

mod security_and_parts;
mod security_indicators;
mod threat_indicators;

fn build_odf_zip(mimetype: &str, content_xml: &str, styles_xml: Option<&str>) -> Vec<u8> {
    build_odf_zip_custom(mimetype, content_xml, styles_xml, None, Vec::new())
}

#[derive(Default)]
struct RichContentNodeCounts {
    table: usize,
    comment: usize,
    footnote: usize,
    bookmark: usize,
    field: usize,
    shape: usize,
    revision: usize,
    styles: usize,
}

fn collect_rich_content_counts(parsed: &crate::parser::ParsedDocument) -> RichContentNodeCounts {
    let mut counts = RichContentNodeCounts::default();
    for node in parsed.store.values() {
        match node {
            IRNode::Table(_) => counts.table += 1,
            IRNode::Comment(_) => counts.comment += 1,
            IRNode::Footnote(_) => counts.footnote += 1,
            IRNode::BookmarkStart(_) | IRNode::BookmarkEnd(_) => counts.bookmark += 1,
            IRNode::Field(_) => counts.field += 1,
            IRNode::Shape(_) => counts.shape += 1,
            IRNode::Revision(_) => counts.revision += 1,
            IRNode::StyleSet(_) => counts.styles += 1,
            _ => {}
        }
    }
    counts
}

fn build_odf_zip_custom(
    mimetype: &str,
    content_xml: &str,
    styles_xml: Option<&str>,
    manifest_xml: Option<&str>,
    extra_files: Vec<(&str, &[u8])>,
) -> Vec<u8> {
    let mut buffer = Vec::new();
    let cursor = Cursor::new(&mut buffer);
    let mut zip = ZipWriter::new(cursor);
    let stored = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(mimetype.as_bytes()).unwrap();

    zip.start_file("META-INF/manifest.xml", FileOptions::<()>::default())
        .unwrap();
    if let Some(xml) = manifest_xml {
        zip.write_all(xml.as_bytes()).unwrap();
    } else {
        zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#,
            )
            .unwrap();
    }

    zip.start_file("content.xml", FileOptions::<()>::default())
        .unwrap();
    zip.write_all(content_xml.as_bytes()).unwrap();

    zip.start_file("meta.xml", FileOptions::<()>::default())
        .unwrap();
    zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
  <office:meta>
    <dc:title>Test Doc</dc:title>
    <dc:creator>docir</dc:creator>
  </office:meta>
</office:document-meta>
"#,
        )
        .unwrap();

    if let Some(styles) = styles_xml {
        zip.start_file("styles.xml", FileOptions::<()>::default())
            .unwrap();
        zip.write_all(styles.as_bytes()).unwrap();
    }

    for (path, bytes) in extra_files {
        zip.start_file(path, FileOptions::<()>::default()).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap();
    buffer
}

#[test]
fn test_parse_odt_minimal() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:p>Hello ODF</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();
    let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    assert_eq!(parsed.format, DocumentFormat::OdfText);
    let doc = parsed.document().unwrap();
    assert!(!doc.content.is_empty());
}

#[test]
fn test_parse_odt_reports_malformed_list_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:list text:style-name="L1" text:style-name="L2">
        <text:list-item><text:p>Item</text:p></text:list-item>
      </text:list>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed list attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_odt_reports_malformed_note_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:note text:note-class="footnote" text:note-class="endnote">
        <text:note-body><text:p>Note</text:p></text:note-body>
      </text:note>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed note attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_odt_reports_malformed_bookmark_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:bookmark-start text:name="bm1" text:name="bm2"/>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed bookmark attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_odt_reports_malformed_outline_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    for content_xml in [
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:h text:outline-level="1" text:outline-level="2">Heading</text:h>
    </office:text>
  </office:body>
</office:document-content>
"#,
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:h text:outline-level="bad">Heading</text:h>
    </office:text>
  </office:body>
</office:document-content>
"#,
    ] {
        let zip_data = build_odf_zip(mimetype, content_xml, None);
        let parser = DocumentParser::new();

        let err = parser
            .parse_reader(Cursor::new(zip_data))
            .expect_err("malformed outline attributes must fail");

        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }
}

#[test]
fn test_parse_odt_reports_malformed_inline_space_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    for content_xml in [
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:p>Before<text:s text:c="1" text:c="2"/>After</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#,
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:p>Before<text:s text:c="bad"/>After</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#,
    ] {
        let zip_data = build_odf_zip(mimetype, content_xml, None);
        let parser = DocumentParser::new();

        let err = parser
            .parse_reader(Cursor::new(zip_data))
            .expect_err("malformed inline space attributes must fail");

        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }
}

#[test]
fn test_parse_odt_reports_malformed_table_cell_span_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <table:table>
        <table:table-row>
          <table:table-cell table:number-columns-spanned="2" table:number-columns-spanned="3"/>
        </table:table-row>
      </table:table>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed table cell span attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_odt_reports_malformed_table_cell_numeric_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <table:table>
        <table:table-row>
          <table:table-cell table:number-columns-spanned="bad"/>
        </table:table-row>
      </table:table>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed table cell numeric attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_odt_reports_malformed_frame_transform_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <draw:frame svg:x="1cm" svg:x="2cm">
        <draw:image xlink:href="Pictures/image.png"/>
      </draw:frame>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed frame transform attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_odt_reports_malformed_frame_image_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <draw:frame>
        <draw:image xlink:href="Pictures/one.png" xlink:href="Pictures/two.png"/>
      </draw:frame>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed frame image attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_ods_minimal() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1" />
      <table:table table:name="Sheet2" />
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    assert_eq!(parsed.format, DocumentFormat::OdfSpreadsheet);
    let doc = parsed.document().unwrap();
    assert_eq!(doc.content.len(), 2);
}

#[test]
fn test_parse_ods_cells_and_validations() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:spreadsheet>
      <table:content-validations>
        <table:content-validation table:name="val1" table:condition="cell-content-is-between(1,10)" table:allow-empty-cell="true" />
      </table:content-validations>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell table:cell-value-type="float" table:cell-value="3.14" />
          <table:table-cell table:cell-value-type="string">
            <text:p>Hello</text:p>
          </table:table-cell>
          <table:table-cell table:formula="of:=SUM([.A1];[.B1])" table:cell-value-type="float" table:cell-value="6.28" table:content-validation-name="val1" />
        </table:table-row>
        <table:table-row table:number-rows-repeated="2">
          <table:table-cell table:cell-value-type="boolean" table:cell-value="true" table:number-columns-repeated="2" />
        </table:table-row>
        <draw:frame draw:name="Chart1">
          <chart:chart />
        </draw:frame>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();
    let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    let doc = parsed.document().unwrap();

    let mut cell_count = 0;
    let mut cell_refs = Vec::new();
    let mut validation_count = 0;
    let mut drawing_count = 0;
    for node in parsed.store.values() {
        match node {
            IRNode::Cell(cell) => {
                cell_count += 1;
                cell_refs.push(cell.reference.clone());
            }
            IRNode::DataValidation(_) => validation_count += 1,
            IRNode::WorksheetDrawing(_) => drawing_count += 1,
            _ => {}
        }
    }
    assert!(cell_count >= 5, "cells: {:?}", cell_refs);
    assert_eq!(validation_count, 1);
    assert_eq!(drawing_count, 1);
    assert_eq!(doc.content.len(), 1);
}

#[test]
fn test_parse_ods_reports_malformed_cell_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell table:cell-value-type="float" table:cell-value="1" table:cell-value="2"/>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed cell attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_ods_reports_malformed_cell_numeric_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell table:number-columns-repeated="bad"/>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed cell numeric attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_ods_reports_malformed_table_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1" table:name="Sheet2"/>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed table attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_ods_reports_malformed_validation_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:content-validations>
        <table:content-validation table:name="val1" table:name="val2"/>
      </table:content-validations>
      <table:table table:name="Sheet1"/>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed validation attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_ods_reports_malformed_conditional_formatting_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    for content_xml in [
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">
        <table:conditional-formatting table:target-range-address="Sheet1.A1" table:target-range-address="Sheet1.B1"/>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#,
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">
        <table:conditional-formatting table:target-range-address="Sheet1.A1">
          <table:conditional-format table:priority="bad" table:condition="cell-content-is-greater-than(1)"/>
        </table:conditional-formatting>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#,
    ] {
        let zip_data = build_odf_zip(mimetype, content_xml, None);
        let parser = DocumentParser::new();

        let err = parser
            .parse_reader(Cursor::new(zip_data))
            .expect_err("malformed conditional formatting attributes must fail");

        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }
}

#[test]
fn test_parse_odp_minimal() {
    let mimetype = "application/vnd.oasis.opendocument.presentation";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1" />
      <draw:page draw:name="Slide 2" />
    </office:presentation>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    assert_eq!(parsed.format, DocumentFormat::OdfPresentation);
    let doc = parsed.document().unwrap();
    assert_eq!(doc.content.len(), 2);
}

#[test]
fn test_parse_odp_reports_malformed_slide_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.presentation";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1" draw:name="Slide 2">
        <draw:frame/>
      </draw:page>
    </office:presentation>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed slide attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_odp_reports_malformed_shape_text_space_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.presentation";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1">
        <draw:custom-shape>
          <text:p>Before<text:s text:c="1" text:c="2"/>After</text:p>
        </draw:custom-shape>
      </draw:page>
    </office:presentation>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed shape text space attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_odp_shapes_and_notes() {
    let mimetype = "application/vnd.oasis.opendocument.presentation";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1" presentation:transition-type="fade" presentation:transition-speed="fast">
        <draw:frame draw:name="Title">
          <draw:text-box>
            <text:p>Hello ODP</text:p>
          </draw:text-box>
        </draw:frame>
        <draw:frame draw:name="Image1">
          <draw:image xlink:href="Pictures/img1.png" />
        </draw:frame>
        <draw:frame draw:name="Chart1">
          <chart:chart />
        </draw:frame>
        <presentation:notes>
          <text:p>Speaker note</text:p>
        </presentation:notes>
      </draw:page>
    </office:presentation>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

    let mut shape_count = 0;
    let mut slide_notes = 0;
    let mut transition_count = 0;
    for node in parsed.store.values() {
        if let IRNode::Slide(slide) = node {
            if slide.notes.is_some() {
                slide_notes += 1;
            }
            if slide.transition.is_some() {
                transition_count += 1;
            }
        }
        if let IRNode::Shape(_) = node {
            shape_count += 1;
        }
    }

    assert_eq!(shape_count, 3);
    assert_eq!(slide_notes, 1);
    assert_eq!(transition_count, 1);
}

#[test]
fn test_parse_odt_rich_content() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:dc="http://purl.org/dc/elements/1.1/">
  <office:body>
    <office:text>
      <text:p>Intro</text:p>
      <text:list text:style-name="L1">
        <text:list-item>
          <text:p>Item 1</text:p>
        </text:list-item>
      </text:list>
      <table:table>
        <table:table-row>
          <table:table-cell table:number-columns-spanned="2">
            <text:p>Cell A</text:p>
          </table:table-cell>
        </table:table-row>
      </table:table>
      <office:annotation>
        <dc:creator>Alice</dc:creator>
        <dc:date>2024-01-01</dc:date>
        <text:p>Comment body</text:p>
      </office:annotation>
      <text:note text:note-class="footnote">
        <text:note-body>
          <text:p>Footnote body</text:p>
        </text:note-body>
      </text:note>
      <text:bookmark-start text:name="bm1" />
      <text:bookmark-end text:name="bm1" />
      <text:date />
      <draw:frame draw:name="Image1">
        <draw:image xlink:href="Pictures/image1.png" />
      </draw:frame>
      <text:tracked-changes>
        <text:changed-region>
          <text:change-info>
            <dc:creator>Bob</dc:creator>
            <dc:date>2024-01-02</dc:date>
          </text:change-info>
          <text:insertion>
            <text:p>Inserted text</text:p>
          </text:insertion>
        </text:changed-region>
      </text:tracked-changes>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
  <office:styles>
    <style:style style:name="P1" style:family="paragraph" />
  </office:styles>
</office:document-styles>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, Some(styles_xml));
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    let counts = collect_rich_content_counts(&parsed);

    assert_eq!(counts.table, 1);
    assert_eq!(counts.comment, 1);
    assert_eq!(counts.footnote, 1);
    assert!(counts.bookmark >= 2);
    assert_eq!(counts.field, 1);
    assert_eq!(counts.shape, 1);
    assert_eq!(counts.revision, 1);
    assert_eq!(counts.styles, 1);
}

#[test]
fn test_parse_odt_text_content_accepts_alternate_namespace_prefixes() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<pkg:document-content xmlns:pkg="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:off="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:txt="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:tbl="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:drw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:dct="http://purl.org/dc/elements/1.1/">
  <off:body>
    <off:text>
      <txt:h txt:outline-level="2">Heading</txt:h>
      <txt:p>First<txt:s txt:c="2"/>line<txt:tab/>tab<txt:line-break/>next</txt:p>
      <txt:list txt:style-name="AltList">
        <txt:list-item><txt:p>Item</txt:p></txt:list-item>
      </txt:list>
      <tbl:table>
        <tbl:table-row>
          <tbl:table-cell><txt:p>Cell</txt:p></tbl:table-cell>
        </tbl:table-row>
      </tbl:table>
      <off:annotation>
        <dct:creator>Alice</dct:creator>
        <dct:date>2024-01-01</dct:date>
        <txt:p>Comment body</txt:p>
      </off:annotation>
      <txt:note txt:note-class="footnote">
        <txt:note-body><txt:p>Footnote body</txt:p></txt:note-body>
      </txt:note>
      <txt:bookmark-start txt:name="bm-alt" />
      <txt:bookmark-end txt:name="bm-alt" />
      <txt:date />
      <txt:time />
      <drw:frame drw:name="AltImage">
        <drw:image xlink:href="Pictures/image1.png" />
      </drw:frame>
      <txt:tracked-changes>
        <txt:changed-region>
          <txt:change-info>
            <dct:creator>Bob</dct:creator>
            <dct:date>2024-01-02</dct:date>
          </txt:change-info>
          <txt:insertion><txt:p>Inserted text</txt:p></txt:insertion>
        </txt:changed-region>
      </txt:tracked-changes>
    </off:text>
  </off:body>
</pkg:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    let counts = collect_rich_content_counts(&parsed);

    let paragraph_texts = parsed
        .store
        .values()
        .filter_map(|node| match node {
            IRNode::Paragraph(paragraph) => Some(paragraph.text_content(&parsed.store)),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert!(
        paragraph_texts
            .iter()
            .any(|text| text == "First  line\ttab\nnext")
    );
    assert!(parsed.store.values().any(|node| {
        matches!(node, IRNode::Paragraph(paragraph) if paragraph.properties.outline_level == Some(2))
    }));
    assert!(parsed.store.values().any(|node| {
        match node {
            IRNode::Paragraph(paragraph) => paragraph
                .properties
                .numbering
                .as_ref()
                .map(|numbering| numbering.level == 0)
                .unwrap_or(false),
            _ => false,
        }
    }));
    assert_eq!(counts.table, 1);
    assert_eq!(counts.comment, 1);
    assert_eq!(counts.footnote, 1);
    assert!(counts.bookmark >= 2);
    assert_eq!(counts.field, 2);
    assert_eq!(counts.shape, 1);
    assert_eq!(counts.revision, 1);
    assert!(parsed.store.values().any(|node| {
        matches!(
            node,
            IRNode::Shape(shape)
                if shape.name.as_deref() == Some("AltImage")
                    && shape.media_target.as_deref() == Some("Pictures/image1.png")
        )
    }));
}
