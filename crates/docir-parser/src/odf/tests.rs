use super::*;
use crate::parser::DocumentParser;
use docir_core::security::ThreatIndicatorType;
use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

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
fn test_parse_odf_security_indicators() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = odf_security_content_xml();
    let manifest_xml = odf_security_manifest_xml();
    let signatures_xml = odf_security_signatures_xml();
    let zip_data = build_odf_zip_custom(
        mimetype,
        &content_xml,
        None,
        Some(&manifest_xml),
        vec![
            ("META-INF/documentsignatures.xml", &signatures_xml),
            ("Scripts/macro.py", b"print('hi')"),
            ("Object 1", b"oledata"),
        ],
    );
    let parser = DocumentParser::new();
    let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    assert_security_content_present(&parsed);
    assert_signature_and_encryption_diagnostics(&parsed);
}

fn odf_security_content_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <text:p>
        <text:a xlink:href="https://example.com">Link</text:a>
      </text:p>
      <draw:object-ole xlink:href="https://example.com/ole.bin" />
    </office:text>
  </office:body>
  <office:scripts>
    <script:script xlink:href="Scripts/macro.py" />
  </office:scripts>
</office:document-content>
"#
    .to_string()
}

fn odf_security_manifest_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Scripts/macro.py" manifest:media-type="application/vnd.sun.star.script"/>
  <manifest:file-entry manifest:full-path="Object 1" manifest:media-type="application/vnd.sun.star.oleobject"/>
  <manifest:file-entry manifest:full-path="Encrypted" manifest:media-type="application/vnd.oasis.opendocument.encrypted"/>
</manifest:manifest>
"#
    .to_string()
}

fn odf_security_signatures_xml() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8"?>
<ds:Signatures xmlns:ds="http://www.w3.org/2000/09/xmldsig#">
  <ds:Signature>
    <ds:SignedInfo>
      <ds:SignatureMethod Algorithm="http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"/>
      <ds:Reference>
        <ds:DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>
      </ds:Reference>
    </ds:SignedInfo>
    <ds:KeyInfo>
      <ds:X509Data>
        <ds:X509SubjectName>CN=Tester</ds:X509SubjectName>
      </ds:X509Data>
    </ds:KeyInfo>
  </ds:Signature>
</ds:Signatures>
"#
    .to_vec()
}

fn assert_security_content_present(parsed: &ParsedDocument) {
    let doc = parsed.document().unwrap();
    assert!(doc.security.macro_project.is_some());
    assert!(!doc.security.external_refs.is_empty());
    assert!(!doc.security.ole_objects.is_empty());
}

fn assert_signature_and_encryption_diagnostics(parsed: &ParsedDocument) {
    let mut sig_count = 0;
    let mut encryption_diag = false;
    for node in parsed.store.values() {
        if let IRNode::DigitalSignature(_) = node {
            sig_count += 1;
        }
        if let IRNode::Diagnostics(diag) = node {
            if diag.entries.iter().any(|e| e.code == "ODF_ENCRYPTION") {
                encryption_diag = true;
            }
        }
    }
    assert_eq!(sig_count, 1);
    assert!(encryption_diag);
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
fn test_odf_formula_dde_and_links() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell table:formula="of:=DDE(&quot;soffice&quot;;&quot;file:///tmp/test.ods&quot;;&quot;A1&quot;)" />
          <table:table-cell table:formula="of:=HYPERLINK(&quot;https://example.com&quot;;&quot;Example&quot;)" />
          <table:table-cell table:formula="of:=WEBSERVICE(&quot;https://example.com/api&quot;)" />
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
    let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#;
    let zip_data =
        build_odf_zip_custom(mimetype, content_xml, None, Some(manifest_xml), Vec::new());
    let parser = DocumentParser::new();
    let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    let doc = parsed.document().unwrap();

    assert!(!doc.security.dde_fields.is_empty());
    assert!(doc
        .security
        .threat_indicators
        .iter()
        .any(|i| i.indicator_type == ThreatIndicatorType::DdeCommand));

    let mut has_formula_link = false;
    let mut has_unsupported = false;
    for node in parsed.store.values() {
        if let IRNode::ExternalReference(ext) = node {
            if ext.target.contains("example.com") {
                has_formula_link = true;
            }
        }
        if let IRNode::Diagnostics(diag) = node {
            if diag
                .entries
                .iter()
                .any(|e| e.code == "ODF_FORMULA_UNSUPPORTED_FUNCTION")
            {
                has_unsupported = true;
            }
        }
    }
    assert!(has_formula_link);
    assert!(has_unsupported);
}

#[test]
fn test_odf_encryption_metadata_diagnostics() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:text><text:p>Encrypted</text:p></office:text></office:body>
</office:document-content>
"#;
    let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml">
    <manifest:encryption-data manifest:checksum-type="SHA1" manifest:checksum="YWJjZA==">
      <manifest:algorithm manifest:algorithm-name="http://www.w3.org/2001/04/xmlenc#aes256-cbc"
        manifest:initialisation-vector="MTIzNDU2Nzg5MA==" manifest:key-size="32"/>
      <manifest:key-derivation manifest:key-derivation-name="PBKDF2"
        manifest:salt="c2FsdA==" manifest:iteration-count="2048"/>
    </manifest:encryption-data>
  </manifest:file-entry>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#;
    let zip_data =
        build_odf_zip_custom(mimetype, content_xml, None, Some(manifest_xml), Vec::new());
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

    let mut has_meta = false;
    for node in parsed.store.values() {
        if let IRNode::Diagnostics(diag) = node {
            if diag.entries.iter().any(|e| e.code == "ODF_ENCRYPTION_META") {
                has_meta = true;
            }
        }
    }
    assert!(has_meta);
}

#[test]
fn test_odf_manifest_inventory_and_parts() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
  <office:body>
    <office:text>
      <text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">Hello</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Thumbnails/thumbnail.png" manifest:media-type="image/png"/>
</manifest:manifest>
"#;
    let zip_data = build_odf_zip_custom(
            mimetype,
            content_xml,
            Some(r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#),
            Some(manifest_xml),
            vec![
                (
                    "settings.xml",
                    br#"<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
                ),
                ("Thumbnails/thumbnail.png", b"pngdata"),
            ],
        );
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

    let mut part_paths = Vec::new();
    let mut asset_paths = Vec::new();
    let mut odf_parts = Vec::new();
    for node in parsed.store.values() {
        match node {
            IRNode::ExtensionPart(part) => part_paths.push(part.path.clone()),
            IRNode::MediaAsset(asset) => asset_paths.push(asset.path.clone()),
            IRNode::Diagnostics(diag) => {
                for entry in &diag.entries {
                    if entry.code == "ODF_PART" {
                        if let Some(path) = entry.path.as_ref() {
                            odf_parts.push(path.clone());
                        }
                    }
                }
            }
            _ => {}
        }
    }

    assert!(part_paths.contains(&"content.xml".to_string()));
    assert!(part_paths.contains(&"styles.xml".to_string()));
    assert!(part_paths.contains(&"settings.xml".to_string()));
    assert!(asset_paths.contains(&"Thumbnails/thumbnail.png".to_string()));
    assert!(odf_parts.contains(&"content.xml".to_string()));
    assert!(odf_parts.contains(&"styles.xml".to_string()));
}

#[test]
fn test_odt_headers_and_footers_from_styles() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:p>Body</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:master-styles>
    <style:master-page style:name="Standard">
      <style:header>
        <text:p>Header text</text:p>
      </style:header>
      <style:footer>
        <text:p>Footer text</text:p>
      </style:footer>
    </style:master-page>
  </office:master-styles>
</office:document-styles>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, Some(styles_xml));
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

    let mut header_texts = Vec::new();
    let mut footer_texts = Vec::new();
    for node in parsed.store.values() {
        match node {
            IRNode::Header(header) => {
                for id in &header.content {
                    if let Some(IRNode::Paragraph(p)) = parsed.store.get(*id) {
                        for run_id in &p.runs {
                            if let Some(IRNode::Run(run)) = parsed.store.get(*run_id) {
                                header_texts.push(run.text.clone());
                            }
                        }
                    }
                }
            }
            IRNode::Footer(footer) => {
                for id in &footer.content {
                    if let Some(IRNode::Paragraph(p)) = parsed.store.get(*id) {
                        for run_id in &p.runs {
                            if let Some(IRNode::Run(run)) = parsed.store.get(*run_id) {
                                footer_texts.push(run.text.clone());
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    assert!(header_texts.iter().any(|t| t == "Header text"));
    assert!(footer_texts.iter().any(|t| t == "Footer text"));
}

#[test]
fn test_parse_ods_named_ranges_and_pivots() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:spreadsheet>
      <table:named-expressions>
        <table:named-range table:name="RANGE1" table:cell-range-address="Sheet1.A1:Sheet1.B2"/>
        <table:named-expression table:name="EXPR1" table:expression="of:=SUM([.A1];[.B1])"/>
      </table:named-expressions>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell table:cell-value-type="float" table:cell-value="1"/>
          <table:table-cell table:cell-value-type="float" table:cell-value="2"/>
        </table:table-row>
      </table:table>
      <table:data-pilot-table table:name="Pivot1"
        table:source-range-address="Sheet1.A1:Sheet1.B2"
        table:target-range-address="Sheet1.D1:Sheet1.E2">
        <table:data-pilot-field table:source-field-name="Field1"/>
      </table:data-pilot-table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();
    let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    let doc = parsed.document().unwrap();

    assert!(!doc.defined_names.is_empty());

    let mut pivot_tables = 0;
    let mut pivot_caches = 0;
    let mut pivot_records = 0;
    for node in parsed.store.values() {
        match node {
            IRNode::PivotTable(_) => pivot_tables += 1,
            IRNode::PivotCache(_) => pivot_caches += 1,
            IRNode::PivotCacheRecords(_) => pivot_records += 1,
            _ => {}
        }
    }

    assert!(pivot_tables >= 1);
    assert!(pivot_caches >= 1);
    assert!(pivot_records >= 1);
}

#[test]
fn test_parse_odp_master_pages_and_transitions() {
    let mimetype = "application/vnd.oasis.opendocument.presentation";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1" presentation:transition-type="fade" presentation:transition-speed="fast"/>
    </office:presentation>
  </office:body>
</office:document-content>
"#;
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
  <office:master-styles>
    <style:master-page style:name="Master1"/>
  </office:master-styles>
</office:document-styles>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, Some(styles_xml));
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

    let mut slide_with_transition = 0;
    let mut master_page_diag = false;
    for node in parsed.store.values() {
        if let IRNode::Slide(slide) = node {
            if slide.transition.is_some() {
                slide_with_transition += 1;
            }
        }
        if let IRNode::Diagnostics(diag) = node {
            if diag.entries.iter().any(|e| e.code == "ODF_MASTER_PAGE") {
                master_page_diag = true;
            }
        }
    }

    assert_eq!(slide_with_transition, 1);
    assert!(master_page_diag);
}
