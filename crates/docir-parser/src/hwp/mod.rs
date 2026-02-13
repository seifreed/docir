//! HWP/HWPX parsing (Hangul Word Processor).

use crate::diagnostics::{attach_diagnostics_if_any, push_info, push_warning};
use crate::error::ParseError;
use crate::format::FormatParser;
use crate::input::{enforce_input_size, read_all_with_limit};
use crate::ole::Cfb;
use crate::parser::{ParsedDocument, ParserConfig};
use crate::text_utils::parse_text_alignment;
use crate::xml_utils::{attr_value_by_suffix, local_name};
use crate::zip_handler::SecureZipReader;
use docir_core::ir::{
    Comment, CommentReference, Diagnostics, Document, Endnote, ExtensionPart, ExtensionPartKind,
    Footer, Footnote, Header, IRNode, MediaAsset, MediaType, NumberingInfo, Paragraph, Revision,
    RevisionType, Run, RunProperties, Section, Shape, ShapeType, Style, StyleParagraphProperties,
    StyleRunProperties, StyleSet, StyleType, Table, TableAlignment, TableCell, TableProperties,
    TableRow, TableWidth, TableWidthType,
};
use docir_core::security::{
    ExternalRefType, ExternalReference, MacroModule, MacroModuleType, MacroProject, OleObject,
};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use io::{dump_hwp_streams, prepare_hwp_stream_data};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

mod builder;
mod io;
mod legacy;
mod styles;
use std::collections::HashMap;
use std::io::{Read, Seek};

pub mod part_registry;
mod section;

use legacy::{
    maybe_decompress_stream, parse_default_jscript, parse_docinfo_section_count, parse_file_header,
    parse_hwp_section_stream, scan_hwp_external_refs,
};
use section::parse_hwpx_section;

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

impl HwpParser {
    pub fn new() -> Self {
        Self {
            config: ParserConfig::default(),
        }
    }

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

impl HwpxParser {
    pub fn new() -> Self {
        Self {
            config: ParserConfig::default(),
        }
    }

    pub fn with_config(config: ParserConfig) -> Self {
        Self { config }
    }

    crate::impl_parse_entrypoints!();
}

fn is_hwpx_section(path: &str) -> bool {
    path.starts_with("Contents/section") && path.ends_with(".xml")
}

fn is_hwpx_header(path: &str) -> bool {
    path.starts_with("Contents/header") && path.ends_with(".xml")
}

fn is_hwpx_footer(path: &str) -> bool {
    path.starts_with("Contents/footer") && path.ends_with(".xml")
}

fn is_hwpx_master(path: &str) -> bool {
    path.starts_with("Contents/masterPage") && path.ends_with(".xml")
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

fn attr_any(e: &BytesStart, names: &[&[u8]]) -> Option<String> {
    for name in names {
        for attr in e.attributes().flatten() {
            if attr.key.as_ref() == *name {
                if let Ok(value) = attr.unescape_value() {
                    return Some(value.to_string());
                }
            }
        }
    }
    None
}

fn run_properties_from_attrs(e: &BytesStart) -> RunProperties {
    let mut props = RunProperties::default();
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        let Ok(value) = attr.unescape_value() else {
            continue;
        };
        let value = value.to_string();
        match key {
            b"bold" | b"b" => {
                props.bold = Some(value == "1" || value.eq_ignore_ascii_case("true"));
            }
            b"italic" | b"i" => {
                props.italic = Some(value == "1" || value.eq_ignore_ascii_case("true"));
            }
            b"underline" | b"u" => {
                props.underline = Some(docir_core::ir::UnderlineStyle::Single);
            }
            b"color" => {
                props.color = Some(value.trim_start_matches('#').to_string());
            }
            b"highlight" => {
                props.highlight = Some(value.trim_start_matches('#').to_string());
            }
            b"font" | b"fontName" => {
                props.font_family = Some(value);
            }
            b"size" | b"fontSize" => {
                if let Ok(size) = value.parse::<u32>() {
                    props.font_size = Some(size);
                }
            }
            _ => {}
        }
    }
    props
}

fn style_run_props_from_run(run: RunProperties) -> StyleRunProperties {
    StyleRunProperties {
        font_family: run.font_family,
        font_size: run.font_size,
        bold: run.bold,
        italic: run.italic,
        underline: run.underline,
        strike: run.strike,
        color: run.color,
        highlight: run.highlight,
        vertical_align: run.vertical_align,
        all_caps: run.all_caps,
        small_caps: run.small_caps,
    }
}

fn parse_hwpx_paragraph_props(e: &BytesStart) -> StyleParagraphProperties {
    let mut props = StyleParagraphProperties::default();
    if let Some(align) = attr_any(e, &[b"align", b"alignment", b"textAlign"]) {
        props.alignment = parse_text_alignment(&align);
    }
    let mut indent = docir_core::ir::Indentation::default();
    let mut has_indent = false;
    if let Some(value) = attr_any(e, &[b"indentLeft", b"indent-left", b"left"]) {
        if let Ok(left) = value.parse::<i32>() {
            indent.left = Some(left);
            has_indent = true;
        }
    }
    if let Some(value) = attr_any(e, &[b"indentRight", b"indent-right", b"right"]) {
        if let Ok(right) = value.parse::<i32>() {
            indent.right = Some(right);
            has_indent = true;
        }
    }
    if let Some(value) = attr_any(e, &[b"firstIndent", b"first-indent", b"first"]) {
        if let Ok(first) = value.parse::<i32>() {
            indent.first_line = Some(first);
            has_indent = true;
        }
    }
    if has_indent {
        props.indentation = Some(indent);
    }
    props
}

fn parse_hwpx_table_props(e: &BytesStart) -> Option<TableProperties> {
    let mut props = TableProperties::default();
    let mut has_value = false;
    if let Some(width) = attr_any(e, &[b"width", b"w", b"tableWidth"]) {
        if let Ok(value) = width.parse::<u32>() {
            props.width = Some(TableWidth {
                value,
                width_type: TableWidthType::Dxa,
            });
            has_value = true;
        }
    }
    if let Some(align) = attr_any(e, &[b"align", b"alignment", b"tableAlign"]) {
        let align = align.to_ascii_lowercase();
        props.alignment = match align.as_str() {
            "left" => Some(TableAlignment::Left),
            "center" => Some(TableAlignment::Center),
            "right" => Some(TableAlignment::Right),
            _ => None,
        };
        if props.alignment.is_some() {
            has_value = true;
        }
    }
    if has_value {
        Some(props)
    } else {
        None
    }
}

fn media_type_from_path(path: &str) -> MediaType {
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".png")
        || lower.ends_with(".jpg")
        || lower.ends_with(".jpeg")
        || lower.ends_with(".gif")
        || lower.ends_with(".bmp")
    {
        MediaType::Image
    } else if lower.ends_with(".mp3") || lower.ends_with(".wav") || lower.ends_with(".aac") {
        MediaType::Audio
    } else if lower.ends_with(".mp4")
        || lower.ends_with(".avi")
        || lower.ends_with(".mov")
        || lower.ends_with(".wmv")
    {
        MediaType::Video
    } else {
        MediaType::Other
    }
}

struct HwpxSecurityScan {
    external_refs: Vec<ExternalReference>,
    macro_modules: Vec<NodeId>,
    has_autoexec: bool,
    encrypted_flag: bool,
    protected_flag: bool,
}

impl HwpxSecurityScan {
    fn new() -> Self {
        Self {
            external_refs: Vec::new(),
            macro_modules: Vec::new(),
            has_autoexec: false,
            encrypted_flag: false,
            protected_flag: false,
        }
    }

    fn mark_flags_from_text(&mut self, text: &str) {
        let lower = text.to_ascii_lowercase();
        if lower.contains("encrypt") {
            self.encrypted_flag = true;
        }
        if lower.contains("protect") || lower.contains("password") {
            self.protected_flag = true;
        }
    }
}

fn scan_hwpx_security<R: Read + Seek>(
    file_names: &[String],
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
    doc: &mut Document,
) {
    let mut scan = HwpxSecurityScan::new();
    let mut diagnostics = Diagnostics::new();
    diagnostics.span = Some(SourceSpan::new("package"));

    for path in file_names {
        scan_hwpx_security_path(path, zip, store, &mut scan);
    }

    for ext in scan.external_refs {
        store.insert(IRNode::ExternalReference(ext));
    }
    if !scan.macro_modules.is_empty() {
        let mut project = MacroProject::new();
        project.name = Some("HWPX Scripts".to_string());
        project.modules = scan.macro_modules;
        project.has_auto_exec = scan.has_autoexec;
        project.span = Some(SourceSpan::new("package"));
        store.insert(IRNode::MacroProject(project));
    }
    if scan.encrypted_flag {
        push_warning(
            &mut diagnostics,
            "HWPX_ENCRYPTED",
            "HWPX encrypted content detected".to_string(),
            None,
        );
    }
    if scan.protected_flag {
        push_info(
            &mut diagnostics,
            "HWPX_PROTECTED",
            "HWPX protected content detected".to_string(),
            None,
        );
    }
    attach_diagnostics_if_any(store, doc, diagnostics);
}

fn scan_hwpx_security_path<R: Read + Seek>(
    path: &str,
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
    scan: &mut HwpxSecurityScan,
) {
    let lower = path.to_ascii_lowercase();
    scan.mark_flags_from_text(&lower);
    scan_hwpx_binary_ole(path, &lower, zip, store);
    scan_hwpx_script(path, &lower, zip, store, scan);
    if path.ends_with(".xml") {
        scan_hwpx_xml(path, zip, store, scan);
    }
}

fn scan_hwpx_binary_ole<R: Read + Seek>(
    path: &str,
    lower: &str,
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
) {
    if !lower.starts_with("bindata/") {
        return;
    }
    if !(lower.contains("ole") || lower.contains("object")) {
        return;
    }
    let mut ole = OleObject::new();
    ole.name = Some(path.to_string());
    ole.size_bytes = zip.file_size(path).unwrap_or(0);
    store.insert(IRNode::OleObject(ole));
}

fn scan_hwpx_script<R: Read + Seek>(
    path: &str,
    lower: &str,
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
    scan: &mut HwpxSecurityScan,
) {
    let is_script = lower.starts_with("scripts/")
        || lower.contains("/scripts/")
        || lower.starts_with("macros/")
        || lower.contains("/macros/")
        || lower.ends_with(".js")
        || lower.ends_with(".vbs")
        || lower.ends_with(".wsf")
        || lower.ends_with(".sct")
        || lower.ends_with(".py");
    if !is_script {
        return;
    }

    if let Ok(data) = zip.read_file(path) {
        let source = String::from_utf8_lossy(&data).to_string();
        let mut module = MacroModule::new(path.to_string(), MacroModuleType::Standard);
        module.source_code = Some(source.clone());
        module.span = Some(SourceSpan::new(path));
        let id = module.id;
        store.insert(IRNode::MacroModule(module));
        scan.macro_modules.push(id);

        let lower_source = source.to_ascii_lowercase();
        if lower.contains("auto")
            || lower_source.contains("autoexec")
            || lower_source.contains("auto_open")
            || lower_source.contains("onopen")
        {
            scan.has_autoexec = true;
        }
    }
}

fn scan_hwpx_xml<R: Read + Seek>(
    path: &str,
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
    scan: &mut HwpxSecurityScan,
) {
    if let Ok(xml) = zip.read_file_string(path) {
        scan.external_refs
            .extend(scan_hwpx_external_refs(&xml, path));
        scan.mark_flags_from_text(&xml);
        if xml.to_ascii_lowercase().contains("ole") {
            let mut ole = OleObject::new();
            ole.name = Some(path.to_string());
            ole.size_bytes = xml.len() as u64;
            store.insert(IRNode::OleObject(ole));
        }
    }
}

fn scan_hwpx_external_refs(xml: &str, source: &str) -> Vec<ExternalReference> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut refs = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    if key.ends_with(b"href") || key.ends_with(b"src") || key.ends_with(b"link") {
                        if let Ok(value) = attr.unescape_value() {
                            let target = value.to_string();
                            if target.is_empty() {
                                continue;
                            }
                            let mut ext =
                                ExternalReference::new(ExternalRefType::Hyperlink, target);
                            ext.span = Some(SourceSpan::new(source));
                            refs.push(ext);
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    refs
}

fn build_hwp_diagnostics(format: DocumentFormat, paths: &[String]) -> Diagnostics {
    let registry = part_registry::registry_for(format);
    let mut diagnostics = Diagnostics::new();
    diagnostics.span = Some(SourceSpan::new("package"));

    for path in paths {
        push_info(
            &mut diagnostics,
            "HWP_PART",
            format!("part: {}", path),
            Some(path),
        );
    }

    for spec in registry {
        let mut matched = false;
        for path in paths {
            if part_registry::matches_pattern(path, spec.pattern) {
                matched = true;
                break;
            }
        }
        if !matched {
            push_warning(
                &mut diagnostics,
                "COVERAGE_MISSING",
                format!(
                    "missing part for pattern {} (expected parser={})",
                    spec.pattern, spec.expected_parser
                ),
                Some(spec.pattern),
            );
        }
    }

    diagnostics
}

#[cfg(test)]
mod tests {
    use super::{is_hwpx_mimetype, HwpxParser};
    use crate::parser::DocumentParser;
    use docir_core::ir::{IRNode, ShapeType};
    use docir_core::types::DocumentFormat;
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
