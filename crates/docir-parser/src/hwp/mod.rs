//! HWP/HWPX parsing (Hangul Word Processor).

use crate::diagnostics::attach_diagnostics_if_any;
use crate::error::ParseError;
use crate::format::FormatParser;
use crate::parser::{ParsedDocument, ParserConfig};
use crate::xml_utils::{attr_value_by_suffix, local_name};
use docir_core::ir::{IRNode, Shape, ShapeType, Table, TableCell, TableRow};
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::BytesStart;

mod builder;
mod helpers;
mod io;
mod legacy;
mod security;
mod styles;
use std::collections::HashMap;
use std::io::{Read, Seek};

pub mod part_registry;
mod section;

#[cfg(test)]
use helpers::build_hwp_diagnostics;
use helpers::{
    attr_any, parse_hwpx_paragraph_props, parse_hwpx_table_props, run_properties_from_attrs,
    style_run_props_from_run,
};
#[cfg(test)]
use helpers::{
    is_hwpx_footer, is_hwpx_header, is_hwpx_master, is_hwpx_section, media_type_from_path,
};
use legacy::{maybe_decompress_stream, parse_file_header, scan_hwp_external_refs};
use security::scan_hwpx_security;

/// Returns true if the mimetype indicates HWPX.
pub fn is_hwpx_mimetype(value: &str) -> bool {
    let lower = value.trim().to_ascii_lowercase();
    lower.contains("hwp+zip") || lower.contains("hwpx")
}

/// Parser for legacy HWP (OLE/CFB).
pub struct HwpParser {
    #[allow(dead_code)]
    config: ParserConfig,
}

impl FormatParser for HwpParser {
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        self.parse_reader(reader)
    }
}

impl Default for HwpParser {
    fn default() -> Self {
        Self::new()
    }
}

impl HwpParser {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            config: ParserConfig::default(),
        }
    }

    /// Public API entrypoint: with_config.
    pub fn with_config(config: ParserConfig) -> Self {
        Self { config }
    }

    crate::impl_parse_entrypoints!();
}

/// Parser for HWPX (ZIP + XML).
pub struct HwpxParser {
    config: ParserConfig,
}

impl FormatParser for HwpxParser {
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        self.parse_reader(reader)
    }
}

impl Default for HwpxParser {
    fn default() -> Self {
        Self::new()
    }
}

impl HwpxParser {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            config: ParserConfig::default(),
        }
    }

    /// Public API entrypoint: with_config.
    pub fn with_config(config: ParserConfig) -> Self {
        Self { config }
    }

    crate::impl_parse_entrypoints!();
}

fn parse_hwpx_shape(
    e: &BytesStart<'_>,
    local: &[u8],
    source: &str,
    media_lookup: &HashMap<String, NodeId>,
    store: &mut IrStore,
) -> Option<NodeId> {
    let shape_type = match local {
        b"pic" | b"image" | b"img" => ShapeType::Picture,
        b"chart" => ShapeType::Chart,
        b"audio" => ShapeType::Audio,
        b"video" => ShapeType::Video,
        b"ole" | b"object" => ShapeType::OleObject,
        b"shape" | b"draw" => ShapeType::Custom,
        _ => ShapeType::Unknown,
    };
    if matches!(shape_type, ShapeType::Unknown) {
        return None;
    }

    let mut shape = Shape::new(shape_type);
    shape.name = attr_any(e, &[b"name", b"id", b"shapeId", b"shape-id"]);
    shape.alt_text = attr_any(e, &[b"alt", b"altText", b"alt-text"]);
    shape.hyperlink = attr_any(e, &[b"href", b"link", b"xlink:href"]);
    shape.media_target = attr_value_by_suffix(
        e,
        &[
            b"href",
            b"src",
            b"link",
            b"binaryRef",
            b"binData",
            b"binDataId",
        ],
    );
    if let Some(target) = shape.media_target.as_deref() {
        if let Some(id) = media_lookup.get(target) {
            shape.media_asset = Some(*id);
        }
    }
    if let Some(x) = attr_any(e, &[b"x", b"posX", b"left"]).and_then(|v| v.parse::<i64>().ok()) {
        shape.transform.x = x;
    }
    if let Some(y) = attr_any(e, &[b"y", b"posY", b"top"]).and_then(|v| v.parse::<i64>().ok()) {
        shape.transform.y = y;
    }
    if let Some(width) = attr_any(e, &[b"width", b"w"]).and_then(|v| v.parse::<u64>().ok()) {
        shape.transform.width = width;
    }
    if let Some(height) = attr_any(e, &[b"height", b"h"]).and_then(|v| v.parse::<u64>().ok()) {
        shape.transform.height = height;
    }
    shape.span = Some(SourceSpan::new(source));

    let shape_id = shape.id;
    store.insert(IRNode::Shape(shape));
    Some(shape_id)
}

fn finalize_cell_hwpx(
    current_cell: &mut Option<TableCell>,
    current_row: &mut Option<TableRow>,
    store: &mut IrStore,
) {
    if let Some(cell) = current_cell.take() {
        let cell_id = cell.id;
        store.insert(IRNode::TableCell(cell));
        if let Some(row) = current_row.as_mut() {
            row.cells.push(cell_id);
        }
    }
}

fn finalize_row_hwpx(
    current_row: &mut Option<TableRow>,
    current_table: &mut Option<Table>,
    store: &mut IrStore,
) {
    if let Some(row) = current_row.take() {
        let row_id = row.id;
        store.insert(IRNode::TableRow(row));
        if let Some(table) = current_table.as_mut() {
            table.rows.push(row_id);
        }
    }
}

fn finalize_table_hwpx(
    current_table: &mut Option<Table>,
    current_cell: &mut Option<TableCell>,
    content: &mut Vec<NodeId>,
    store: &mut IrStore,
) {
    if current_cell.is_some() {
        finalize_cell_hwpx(current_cell, &mut None, store);
    }
    if let Some(table) = current_table.take() {
        let table_id = table.id;
        store.insert(IRNode::Table(table));
        content.push(table_id);
    }
}

#[cfg(test)]
mod tests {
    use super::{
        attr_any, build_hwp_diagnostics, is_hwpx_footer, is_hwpx_header, is_hwpx_master,
        is_hwpx_mimetype, is_hwpx_section, media_type_from_path, parse_hwpx_paragraph_props,
        parse_hwpx_table_props, run_properties_from_attrs, style_run_props_from_run, HwpxParser,
    };
    use crate::parser::DocumentParser;
    use docir_core::ir::{DiagnosticSeverity, IRNode, ShapeType, TableAlignment, TableWidthType};
    use docir_core::types::DocumentFormat;
    use quick_xml::events::{BytesStart, Event};
    use quick_xml::Reader;
    use std::io::Write;
    use zip::write::FileOptions;
    use zip::CompressionMethod;

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
        let para = start_event(
            r#"<hp:p align="center" indentLeft="10" indentRight="20" firstIndent="30"/>"#,
        );
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
        assert!(diagnostics
            .entries
            .iter()
            .any(|e| e.code == "HWP_PART" && matches!(e.severity, DiagnosticSeverity::Info)));
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
            if let IRNode::Shape(shape) = node {
                if shape.shape_type == ShapeType::Picture {
                    has_image = true;
                }
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
            if let IRNode::StyleSet(set) = node {
                if let Some(style) = set.styles.iter().find(|s| s.style_id == "s1") {
                    assert!(style.run_props.is_some());
                    assert!(style.paragraph_props.is_some());
                    has_style = true;
                }
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
        assert!(doc
            .security
            .threat_indicators
            .iter()
            .any(|i| i.indicator_type == docir_core::security::ThreatIndicatorType::AutoExecMacro));

        let mut has_protected = false;
        for node in parsed.store.values() {
            if let IRNode::Diagnostics(diag) = node {
                if diag.entries.iter().any(|e| e.code == "HWPX_PROTECTED") {
                    has_protected = true;
                }
            }
        }
        assert!(has_protected);
    }
}
