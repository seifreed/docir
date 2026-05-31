use super::helpers::{
    attr_any, build_hwp_diagnostics, is_hwpx_footer, is_hwpx_header, is_hwpx_master,
    is_hwpx_section, media_type_from_path, parse_hwpx_paragraph_props, parse_hwpx_table_props,
    run_properties_from_attrs, style_run_props_from_run,
};
use super::{HwpxParser, is_hwpx_mimetype};
use crate::parser::DocumentParser;
use docir_core::ir::{DiagnosticSeverity, IRNode, ShapeType, TableAlignment, TableWidthType};
use docir_core::types::DocumentFormat;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::io::Write;
use zip::CompressionMethod;
use zip::write::FileOptions;

fn build_hwpx_zip(section_xml: &str) -> Vec<u8> {
    build_hwpx_zip_with_parts(section_xml, None, Vec::new())
}

fn build_hwpx_zip_with_parts(
    section_xml: &str,
    styles_xml: Option<&str>,
    extra_files: Vec<(&str, &[u8])>,
) -> Vec<u8> {
    let mut buffer = std::io::Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(&mut buffer);
    let stored: FileOptions<'_, ()> =
        FileOptions::default().compression_method(CompressionMethod::Stored);

    writer.start_file("mimetype", stored).unwrap();
    writer.write_all(b"application/hwp+zip").unwrap();

    writer.start_file("META-INF/container.xml", stored).unwrap();
    writer
        .write_all(b"<container><rootfiles/></container>")
        .unwrap();

    writer.start_file("version.xml", stored).unwrap();
    writer.write_all(b"<version>1.0</version>").unwrap();

    if let Some(styles_xml) = styles_xml {
        writer.start_file("Contents/content.hpf", stored).unwrap();
        writer.write_all(styles_xml.as_bytes()).unwrap();
    }

    writer.start_file("Contents/section0.xml", stored).unwrap();
    writer.write_all(section_xml.as_bytes()).unwrap();

    for (path, data) in extra_files {
        writer.start_file(path, stored).unwrap();
        writer.write_all(data).unwrap();
    }

    writer.finish().unwrap();
    buffer.into_inner()
}

#[test]
fn test_hwpx_mimetype_detection() {
    assert!(is_hwpx_mimetype("application/hwp+zip"));
    assert!(is_hwpx_mimetype("application/vnd.hancom.hwpx"));
    assert!(!is_hwpx_mimetype("application/vnd.oasis.opendocument.text"));
}

fn start_event(xml: &str) -> BytesStart<'static> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => return e.into_owned(),
            Ok(Event::Eof) => panic!("missing start event"),
            Ok(_) => {}
            Err(err) => panic!("xml read error: {err}"),
        }
        buf.clear();
    }
}

#[test]
fn test_hwpx_path_classifiers() {
    assert!(is_hwpx_section("Contents/section0.xml"));
    assert!(!is_hwpx_section("Contents/section0.txt"));
    assert!(is_hwpx_header("Contents/header1.xml"));
    assert!(is_hwpx_footer("Contents/footer2.xml"));
    assert!(is_hwpx_master("Contents/masterPage0.xml"));
    assert!(!is_hwpx_master("Contents/masterPage0.bin"));
}

#[test]
fn test_media_type_and_attr_helpers() {
    assert!(matches!(
        media_type_from_path("BinData/pic.JPG"),
        docir_core::ir::MediaType::Image
    ));
    assert!(matches!(
        media_type_from_path("BinData/sound.wav"),
        docir_core::ir::MediaType::Audio
    ));
    assert!(matches!(
        media_type_from_path("BinData/movie.MP4"),
        docir_core::ir::MediaType::Video
    ));
    assert!(matches!(
        media_type_from_path("BinData/file.bin"),
        docir_core::ir::MediaType::Other
    ));

    let e = start_event(r#"<hp:run name="shape-1" altText="preview" unknown="x"/>"#);
    assert_eq!(
        attr_any(&e, &[b"missing", b"name"]).as_deref(),
        Some("shape-1")
    );
    assert_eq!(attr_any(&e, &[b"altText"]).as_deref(), Some("preview"));
    assert!(attr_any(&e, &[b"nope"]).is_none());
}

#[test]
fn test_run_and_style_property_helpers() {
    let e = start_event(
        r##"<hp:r bold="1" italic="true" underline="single" color="#AABBCC" highlight="#00FF00" font="Malgun" size="12"/>"##,
    );
    let run = run_properties_from_attrs(&e);
    assert_eq!(run.bold, Some(true));
    assert_eq!(run.italic, Some(true));
    assert!(run.underline.is_some());
    assert_eq!(run.color.as_deref(), Some("AABBCC"));
    assert_eq!(run.highlight.as_deref(), Some("00FF00"));
    assert_eq!(run.font_family.as_deref(), Some("Malgun"));
    assert_eq!(run.font_size, Some(12));

    let style = style_run_props_from_run(run);
    assert_eq!(style.bold, Some(true));
    assert_eq!(style.font_size, Some(12));
    assert_eq!(style.color.as_deref(), Some("AABBCC"));
}

#[test]
fn test_paragraph_and_table_property_helpers() {
    let para =
        start_event(r#"<hp:p align="center" indentLeft="10" indentRight="20" firstIndent="30"/>"#);
    let para_props = parse_hwpx_paragraph_props(&para);
    assert!(para_props.alignment.is_some());
    let indent = para_props.indentation.expect("indentation");
    assert_eq!(indent.left, Some(10));
    assert_eq!(indent.right, Some(20));
    assert_eq!(indent.first_line, Some(30));

    let table = start_event(r#"<hp:tbl width="7200" align="right"/>"#);
    let table_props = parse_hwpx_table_props(&table).expect("table props");
    let width = table_props.width.expect("table width");
    assert_eq!(width.value, 7200);
    assert_eq!(width.width_type, TableWidthType::Dxa);
    assert_eq!(table_props.alignment, Some(TableAlignment::Right));

    let empty = start_event(r#"<hp:tbl align="unknown"/>"#);
    assert!(parse_hwpx_table_props(&empty).is_none());
}

#[test]
fn test_build_hwp_diagnostics_records_parts_and_missing_patterns() {
    let paths = vec![
        "FileHeader".to_string(),
        "DocInfo".to_string(),
        "BodyText/Section0".to_string(),
    ];
    let diagnostics = build_hwp_diagnostics(DocumentFormat::Hwp, &paths);
    assert!(!diagnostics.entries.is_empty());
    assert!(
        diagnostics
            .entries
            .iter()
            .any(|e| e.code == "HWP_PART" && matches!(e.severity, DiagnosticSeverity::Info))
    );
    assert!(diagnostics.entries.iter().any(|e| {
        e.code == "COVERAGE_MISSING" && matches!(e.severity, DiagnosticSeverity::Warning)
    }));
}

#[test]
fn test_parse_hwpx_sections() {
    let xml = r#"<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml">
  <hp:p><hp:t>Hello HWPX</hp:t></hp:p>
  <hp:p><hp:t>Segundo</hp:t></hp:p>
</hp:section>"#;
    let data = build_hwpx_zip(xml);
    let parser = HwpxParser::new();
    let parsed = parser.parse_bytes(&data).expect("hwpx parse");
    assert_eq!(parsed.format, DocumentFormat::Hwpx);
    let doc = parsed.document().expect("doc");
    assert_eq!(doc.content.len(), 1);
}

#[test]
fn test_document_parser_routes_hwpx() {
    let xml = r#"<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml">
  <hp:p><hp:t>Ruta</hp:t></hp:p>
</hp:section>"#;
    let data = build_hwpx_zip(xml);
    let parser = DocumentParser::new();
    let parsed = parser.parse_bytes(&data).expect("hwpx parse");
    assert_eq!(parsed.format, DocumentFormat::Hwpx);
}

#[test]
fn test_hwpx_comments_revisions_images() {
    let xml = r#"<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <hp:p>
    <hp:t>Hola</hp:t>
    <hp:comment id="c1" author="Ana">
      <hp:p><hp:t>Nota</hp:t></hp:p>
    </hp:comment>
    <hp:commentRef ref="c1" />
    <hp:ins author="Bob" date="2024-01-01">
      <hp:t>Insertado</hp:t>
    </hp:ins>
    <hp:del>
      <hp:t>Eliminado</hp:t>
    </hp:del>
    <hp:pic xlink:href="BinData/image1.png" />
  </hp:p>
</hp:section>"#;
    let data = build_hwpx_zip_with_parts(xml, None, vec![("BinData/image1.png", b"img")]);
    let parser = HwpxParser::new();
    let mut parsed = parser.parse_bytes(&data).expect("hwpx parse");
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    let doc = parsed.document().expect("doc");
    assert!(!doc.comments.is_empty());

    let mut has_revision = false;
    let mut has_image = false;
    for node in parsed.store.values() {
        if let IRNode::Revision(_) = node {
            has_revision = true;
        }
        if let IRNode::Shape(shape) = node
            && shape.shape_type == ShapeType::Picture
        {
            has_image = true;
        }
    }
    assert!(has_revision);
    assert!(has_image);
}

#[test]
fn test_hwpx_styles_parsing() {
    let styles_xml = r##"<hp:package xmlns:hp="http://www.hancom.co.kr/hwpml">
  <hp:styles>
    <hp:style id="s1" name="Body" type="paragraph">
      <hp:paraPr align="center" indentLeft="120" />
      <hp:charPr bold="1" italic="true" color="#FF0000" font="Malgun" size="12" />
    </hp:style>
  </hp:styles>
</hp:package>"##;
    let section_xml = r#"<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml">
  <hp:p><hp:t>Texto</hp:t></hp:p>
</hp:section>"#;
    let data = build_hwpx_zip_with_parts(section_xml, Some(styles_xml), Vec::new());
    let parser = HwpxParser::new();
    let mut parsed = parser.parse_bytes(&data).expect("hwpx parse");
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    let doc = parsed.document().expect("doc");
    assert!(doc.styles.is_some());

    let mut has_style = false;
    for node in parsed.store.values() {
        if let IRNode::StyleSet(set) = node
            && let Some(style) = set.styles.iter().find(|s| s.style_id == "s1")
        {
            assert!(style.run_props.is_some());
            assert!(style.paragraph_props.is_some());
            has_style = true;
        }
    }
    assert!(has_style);
}

#[test]
fn test_hwpx_security_signals() {
    let section_xml = r#"<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <hp:p>
    <hp:t>Link</hp:t>
    <hp:pic xlink:href="BinData/oleObject.bin" />
    <hp:a xlink:href="https://example.com">External</hp:a>
  </hp:p>
  <hp:security password="secret" />
</hp:section>"#;
    let data = build_hwpx_zip_with_parts(
        section_xml,
        None,
        vec![
            ("BinData/oleObject.bin", b"oledata"),
            ("Scripts/AutoExec.js", b"AutoExec();"),
        ],
    );
    let parser = HwpxParser::new();
    let mut parsed = parser.parse_bytes(&data).expect("hwpx parse");
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    let doc = parsed.document().expect("doc");
    assert!(doc.security.macro_project.is_some());
    assert!(!doc.security.external_refs.is_empty());
    assert!(!doc.security.ole_objects.is_empty());
    assert!(
        doc.security
            .threat_indicators
            .iter()
            .any(|i| i.indicator_type == docir_core::security::ThreatIndicatorType::AutoExecMacro)
    );

    let mut has_protected = false;
    for node in parsed.store.values() {
        if let IRNode::Diagnostics(diag) = node
            && diag.entries.iter().any(|e| e.code == "HWPX_PROTECTED")
        {
            has_protected = true;
        }
    }
    assert!(has_protected);
}

#[test]
fn test_hwpx_security_reports_malformed_extra_xml() {
    let section_xml = r#"<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml"><hp:p><hp:t>Ok</hp:t></hp:p></hp:section>"#;
    let data = build_hwpx_zip_with_parts(
        section_xml,
        None,
        vec![(
            "Contents/security-extra.xml",
            br#"<extra href="https://example.test"><dangling>"#,
        )],
    );
    let parser = HwpxParser::new();

    let err = parser
        .parse_bytes(&data)
        .expect_err("malformed HWPX XML scanned for security must fail");

    match err {
        crate::error::ParseError::Xml { file, .. } => {
            assert_eq!(file, "Contents/security-extra.xml");
        }
        other => panic!("expected XML error, got {other}"),
    }
}
