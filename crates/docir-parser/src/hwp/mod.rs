//! HWP/HWPX parsing (Hangul Word Processor).

use crate::diagnostics::{push_info, push_warning};
use crate::error::ParseError;
use crate::format::FormatParser;
use crate::input::{enforce_input_size, read_all_with_limit};
use crate::ole::Cfb;
use crate::parser::{ParsedDocument, ParserConfig};
use crate::text_utils::parse_text_alignment;
use crate::xml_utils::{attr_value_by_suffix, local_name};
use crate::zip_handler::SecureZipReader;
use aes::Aes128;
use cbc::cipher::{block_padding::Pkcs7, BlockDecryptMut, KeyIvInit};
use cbc::Decryptor;
use docir_core::ir::{
    Comment, CommentReference, Diagnostics, Document, Endnote, ExtensionPart, ExtensionPartKind,
    Footer, Footnote, Header, IRNode, MediaAsset, MediaType, NumberingInfo, Paragraph, Revision,
    RevisionType, Run, RunProperties, Section, Shape, ShapeType, Style, StyleParagraphProperties,
    StyleRunProperties, StyleSet, StyleType, Table, TableAlignment, TableCell, TableProperties,
    TableRow, TableWidth, TableWidthType,
};
use docir_core::normalize::normalize_store;
use docir_core::security::{
    ExternalRefType, ExternalReference, MacroModule, MacroModuleType, MacroProject, OleObject,
};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use flate2::read::{DeflateDecoder, ZlibDecoder};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use sha1::{Digest as Sha1Digest, Sha1};
use sha2::Sha256;
use std::collections::HashMap;
use std::io::{Read, Seek};

pub mod part_registry;
mod section;

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

    pub fn parse_reader<R: Read + Seek>(
        &self,
        mut reader: R,
    ) -> Result<ParsedDocument, ParseError> {
        let data = read_all_with_limit(reader, self.config.max_input_size)?;

        let cfb = Cfb::parse(data)?;
        let mut store = IrStore::new();
        let mut doc = Document::new(DocumentFormat::Hwp);

        let mut stream_names = cfb.list_streams();
        stream_names.sort();
        let mut shared_parts = Vec::new();
        let mut sections = Vec::new();
        let mut diagnostics = build_hwp_diagnostics(DocumentFormat::Hwp, &stream_names);

        let header_data = cfb
            .read_stream("FileHeader")
            .ok_or_else(|| ParseError::MissingPart("FileHeader".to_string()))?;
        let header = parse_file_header(&header_data)?;
        push_info(
            &mut diagnostics,
            "HWP_HEADER",
            format!(
                "HWP header: version=0x{:08X} flags=0x{:08X}",
                header.version, header.flags
            ),
            Some("FileHeader"),
        );

        let compressed = header.flags & 0x01 != 0;
        let encrypted = header.flags & 0x02 != 0;
        let force_parse = self.config.hwp.force_parse_encrypted;
        let hwp_password = self.config.hwp.password.as_deref();
        let try_raw_encrypted = encrypted && hwp_password.is_none();
        let allow_parse = !encrypted || force_parse || hwp_password.is_some() || try_raw_encrypted;
        if encrypted {
            push_warning(
                &mut diagnostics,
                "HWP_ENCRYPTED",
                "HWP file is encrypted; content parsing skipped".to_string(),
                Some("FileHeader"),
            );
            if force_parse {
                push_warning(
                    &mut diagnostics,
                    "HWP_FORCE_PARSE",
                    "HWP force-parse enabled for encrypted file".to_string(),
                    Some("FileHeader"),
                );
            }
            if hwp_password.is_some() {
                push_info(
                    &mut diagnostics,
                    "HWP_DECRYPT_ATTEMPT",
                    "HWP decryption attempt enabled".to_string(),
                    Some("FileHeader"),
                );
            }
            if try_raw_encrypted {
                push_warning(
                    &mut diagnostics,
                    "HWP_ENCRYPTED_PARTIAL",
                    "HWP encrypted without password; attempting partial parse of readable streams"
                        .to_string(),
                    Some("FileHeader"),
                );
            }
        }

        if self.config.hwp.dump_streams {
            dump_hwp_streams(
                &cfb,
                &stream_names,
                compressed,
                encrypted,
                hwp_password,
                force_parse,
                try_raw_encrypted,
                &mut diagnostics,
            );
        }

        for path in &stream_names {
            let size = cfb.stream_size(path).unwrap_or(0);
            if path.starts_with("BinData/") {
                let mut asset = MediaAsset::new(path, MediaType::Other, size);
                asset.span = Some(SourceSpan::new(path));
                let asset_id = asset.id;
                store.insert(IRNode::MediaAsset(asset));
                shared_parts.push(asset_id);
            } else {
                let mut part = ExtensionPart::new(path, size, ExtensionPartKind::Legacy);
                part.span = Some(SourceSpan::new(path));
                let part_id = part.id;
                store.insert(IRNode::ExtensionPart(part));
                shared_parts.push(part_id);
            }
        }

        let docinfo_data = cfb.read_stream("DocInfo");
        let mut docinfo_section_count = None;
        if let Some(data) = docinfo_data {
            let data = match prepare_hwp_stream_data(
                &data,
                encrypted,
                hwp_password,
                force_parse,
                try_raw_encrypted,
                "DocInfo",
                &mut diagnostics,
            ) {
                Some(bytes) => bytes,
                None => {
                    push_warning(
                        &mut diagnostics,
                        "HWP_DOCINFO_SKIP",
                        "DocInfo skipped due to encryption or decryption failure".to_string(),
                        Some("DocInfo"),
                    );
                    Vec::new()
                }
            };
            match maybe_decompress_stream(&data, compressed, "DocInfo") {
                Ok(bytes) => {
                    docinfo_section_count = parse_docinfo_section_count(&bytes)?;
                }
                Err(err) => {
                    push_warning(
                        &mut diagnostics,
                        "HWP_DECOMPRESS_FAIL",
                        err.to_string(),
                        Some("DocInfo"),
                    );
                }
            }
            if let Some(count) = docinfo_section_count {
                push_info(
                    &mut diagnostics,
                    "HWP_SECTION_COUNT",
                    format!("DocInfo section count: {}", count),
                    Some("DocInfo"),
                );
            }
        }

        if allow_parse {
            for path in &stream_names {
                if path.starts_with("BodyText/Section") {
                    let data = cfb
                        .read_stream(path)
                        .ok_or_else(|| ParseError::MissingPart(path.to_string()))?;
                    let data = match prepare_hwp_stream_data(
                        &data,
                        encrypted,
                        hwp_password,
                        force_parse,
                        try_raw_encrypted,
                        path,
                        &mut diagnostics,
                    ) {
                        Some(bytes) => bytes,
                        None => continue,
                    };
                    let data = match maybe_decompress_stream(&data, compressed, path) {
                        Ok(bytes) => bytes,
                        Err(err) => {
                            push_warning(
                                &mut diagnostics,
                                "HWP_DECOMPRESS_FAIL",
                                err.to_string(),
                                Some(path),
                            );
                            continue;
                        }
                    };
                    let paragraph_ids = parse_hwp_section_stream(&data, path, &mut store)?;
                    let mut section = Section::new();
                    section.name = Some(path.clone());
                    section.content = paragraph_ids;
                    section.span = Some(SourceSpan::new(path));
                    let section_id = section.id;
                    store.insert(IRNode::Section(section));
                    sections.push(section_id);
                }
            }
        }

        if let Some(expected) = docinfo_section_count {
            if expected as usize != sections.len() {
                push_warning(
                    &mut diagnostics,
                    "HWP_SECTION_MISMATCH",
                    format!(
                        "section count mismatch: docinfo={} parsed={}",
                        expected,
                        sections.len()
                    ),
                    Some("DocInfo"),
                );
            }
        }

        if let Some(script_data) = cfb.read_stream("Scripts/DefaultJScript") {
            if let Some(_project_id) =
                parse_default_jscript(&script_data, &mut store, "Scripts/DefaultJScript")
            {
            }
        }

        if allow_parse {
            let externals = scan_hwp_external_refs(
                &cfb,
                &stream_names,
                compressed,
                encrypted,
                hwp_password,
                force_parse,
                try_raw_encrypted,
                &mut diagnostics,
            );
            for ext in externals {
                store.insert(IRNode::ExternalReference(ext));
            }
        }

        for path in &stream_names {
            if path.starts_with("BinData/") {
                let lower = path.to_ascii_lowercase();
                if lower.contains("ole") || lower.contains("object") {
                    let mut ole = OleObject::new();
                    ole.name = Some(path.clone());
                    ole.size_bytes = cfb.stream_size(path).unwrap_or(0);
                    store.insert(IRNode::OleObject(ole));
                }
            }
        }

        doc.shared_parts = shared_parts;
        doc.content = sections;

        if !diagnostics.entries.is_empty() {
            let diag_id = diagnostics.id;
            store.insert(IRNode::Diagnostics(diagnostics));
            doc.diagnostics.push(diag_id);
        }

        let root_id = doc.id;
        store.insert(IRNode::Document(doc));
        normalize_store(&mut store, root_id);

        Ok(ParsedDocument {
            root_id,
            format: DocumentFormat::Hwp,
            store,
            metrics: None,
        })
    }
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

    pub fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        let mut reader = reader;
        enforce_input_size(&mut reader, self.config.max_input_size)?;
        let mut zip = SecureZipReader::new(reader, self.config.zip_config.clone())?;

        let mut store = IrStore::new();
        let mut doc = Document::new(DocumentFormat::Hwpx);

        let mut file_names: Vec<String> = zip.file_names().map(|s| s.to_string()).collect();
        file_names.sort();

        let mut shared_parts = Vec::new();
        let mut media_assets = Vec::new();
        let mut media_lookup: HashMap<String, NodeId> = HashMap::new();
        for path in &file_names {
            let size = zip.file_size(path).unwrap_or(0);
            if path.starts_with("BinData/") {
                let media_type = media_type_from_path(path);
                let mut asset = MediaAsset::new(path, media_type, size);
                asset.span = Some(SourceSpan::new(path));
                let asset_id = asset.id;
                store.insert(IRNode::MediaAsset(asset));
                media_assets.push(asset_id);
                media_lookup.insert(path.clone(), asset_id);
            }
            let mut part = ExtensionPart::new(path, size, ExtensionPartKind::VendorSpecific);
            part.span = Some(SourceSpan::new(path));
            let part_id = part.id;
            store.insert(IRNode::ExtensionPart(part));
            shared_parts.push(part_id);
        }
        doc.shared_parts = shared_parts;
        doc.shared_parts.extend(media_assets);

        let mut section_ids = Vec::new();
        for path in &file_names {
            if is_hwpx_section(path) {
                let xml = zip.read_file_string(path)?;
                let paragraph_ids = parse_hwpx_section(
                    &xml,
                    path,
                    &mut store,
                    &mut doc.comments,
                    &mut doc.footnotes,
                    &mut doc.endnotes,
                    &media_lookup,
                )?;
                let mut section = Section::new();
                section.name = Some(path.clone());
                section.content = paragraph_ids;
                section.span = Some(SourceSpan::new(path));
                let section_id = section.id;
                store.insert(IRNode::Section(section));
                section_ids.push(section_id);
            }
        }
        doc.content = section_ids;

        let mut header_ids = Vec::new();
        let mut footer_ids = Vec::new();
        let mut master_ids = Vec::new();

        for path in &file_names {
            if is_hwpx_header(path) {
                let xml = zip.read_file_string(path)?;
                let content = parse_hwpx_section(
                    &xml,
                    path,
                    &mut store,
                    &mut doc.comments,
                    &mut doc.footnotes,
                    &mut doc.endnotes,
                    &media_lookup,
                )?;
                let mut header = Header::new();
                header.content = content;
                header.span = Some(SourceSpan::new(path));
                let header_id = header.id;
                store.insert(IRNode::Header(header));
                header_ids.push(header_id);
            } else if is_hwpx_footer(path) {
                let xml = zip.read_file_string(path)?;
                let content = parse_hwpx_section(
                    &xml,
                    path,
                    &mut store,
                    &mut doc.comments,
                    &mut doc.footnotes,
                    &mut doc.endnotes,
                    &media_lookup,
                )?;
                let mut footer = Footer::new();
                footer.content = content;
                footer.span = Some(SourceSpan::new(path));
                let footer_id = footer.id;
                store.insert(IRNode::Footer(footer));
                footer_ids.push(footer_id);
            } else if is_hwpx_master(path) {
                let xml = zip.read_file_string(path)?;
                let content = parse_hwpx_section(
                    &xml,
                    path,
                    &mut store,
                    &mut doc.comments,
                    &mut doc.footnotes,
                    &mut doc.endnotes,
                    &media_lookup,
                )?;
                let mut section = Section::new();
                section.name = Some(format!("master:{}", path));
                section.content = content;
                section.span = Some(SourceSpan::new(path));
                let section_id = section.id;
                store.insert(IRNode::Section(section));
                master_ids.push(section_id);
            }
        }

        for section_id in &doc.content {
            if let Some(IRNode::Section(section)) = store.get_mut(*section_id) {
                section.headers.extend(header_ids.iter().copied());
                section.footers.extend(footer_ids.iter().copied());
            }
        }
        doc.shared_parts.extend(master_ids);

        if zip.contains("Contents/content.hpf") {
            if let Ok(xml) = zip.read_file_string("Contents/content.hpf") {
                if let Some(style_set) = parse_hwpx_styles(&xml, "Contents/content.hpf") {
                    let style_id = style_set.id;
                    store.insert(IRNode::StyleSet(style_set));
                    doc.styles = Some(style_id);
                }
            }
        }

        scan_hwpx_security(&file_names, &mut zip, &mut store, &mut doc);

        let diagnostics = build_hwp_diagnostics(DocumentFormat::Hwpx, &file_names);
        if !diagnostics.entries.is_empty() {
            let diag_id = diagnostics.id;
            store.insert(IRNode::Diagnostics(diagnostics));
            doc.diagnostics.push(diag_id);
        }

        let root_id = doc.id;
        store.insert(IRNode::Document(doc));
        normalize_store(&mut store, root_id);

        Ok(ParsedDocument {
            root_id,
            format: DocumentFormat::Hwpx,
            store,
            metrics: None,
        })
    }
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

fn parse_hwpx_styles(xml: &str, source: &str) -> Option<StyleSet> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut styles = Vec::new();
    let mut current: Option<Style> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if local == b"style" {
                    let style_id = attr_any(&e, &[b"id", b"styleId", b"style-id"])
                        .unwrap_or_else(|| "style".to_string());
                    let name = attr_any(&e, &[b"name", b"styleName", b"style-name"]);
                    let style_type = match attr_any(&e, &[b"type", b"styleType"])
                        .as_deref()
                        .map(|v| v.to_ascii_lowercase())
                    {
                        Some(t) if t == "paragraph" => StyleType::Paragraph,
                        Some(t) if t == "character" => StyleType::Character,
                        Some(t) if t == "table" => StyleType::Table,
                        _ => StyleType::Other,
                    };
                    current = Some(Style {
                        style_id,
                        name,
                        style_type,
                        based_on: attr_any(&e, &[b"basedOn", b"based-on"]),
                        next: attr_any(&e, &[b"next", b"next-style"]),
                        is_default: attr_any(&e, &[b"default", b"isDefault"])
                            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
                            .unwrap_or(false),
                        run_props: None,
                        paragraph_props: None,
                        table_props: None,
                    });
                } else if local == b"charPr" || local == b"characterPr" {
                    if let Some(style) = current.as_mut() {
                        let run_props = run_properties_from_attrs(&e);
                        style.run_props = Some(style_run_props_from_run(run_props));
                    }
                } else if local == b"paraPr" || local == b"paragraphPr" {
                    if let Some(style) = current.as_mut() {
                        style.paragraph_props = Some(parse_hwpx_paragraph_props(&e));
                    }
                } else if local == b"tblPr" || local == b"tablePr" {
                    if let Some(style) = current.as_mut() {
                        style.table_props = parse_hwpx_table_props(&e);
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if local == b"style" {
                    let style_id = attr_any(&e, &[b"id", b"styleId", b"style-id"])
                        .unwrap_or_else(|| "style".to_string());
                    let name = attr_any(&e, &[b"name", b"styleName", b"style-name"]);
                    let style_type = match attr_any(&e, &[b"type", b"styleType"])
                        .as_deref()
                        .map(|v| v.to_ascii_lowercase())
                    {
                        Some(t) if t == "paragraph" => StyleType::Paragraph,
                        Some(t) if t == "character" => StyleType::Character,
                        Some(t) if t == "table" => StyleType::Table,
                        _ => StyleType::Other,
                    };
                    styles.push(Style {
                        style_id,
                        name,
                        style_type,
                        based_on: attr_any(&e, &[b"basedOn", b"based-on"]),
                        next: attr_any(&e, &[b"next", b"next-style"]),
                        is_default: attr_any(&e, &[b"default", b"isDefault"])
                            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
                            .unwrap_or(false),
                        run_props: None,
                        paragraph_props: None,
                        table_props: None,
                    });
                } else if local == b"charPr" || local == b"characterPr" {
                    if let Some(style) = current.as_mut() {
                        let run_props = run_properties_from_attrs(&e);
                        style.run_props = Some(style_run_props_from_run(run_props));
                    }
                } else if local == b"paraPr" || local == b"paragraphPr" {
                    if let Some(style) = current.as_mut() {
                        style.paragraph_props = Some(parse_hwpx_paragraph_props(&e));
                    }
                } else if local == b"tblPr" || local == b"tablePr" {
                    if let Some(style) = current.as_mut() {
                        style.table_props = parse_hwpx_table_props(&e);
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if local == b"style" {
                    if let Some(style) = current.take() {
                        styles.push(style);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    if styles.is_empty() {
        None
    } else {
        let mut set = StyleSet::new();
        set.styles = styles;
        set.span = Some(SourceSpan::new(source));
        Some(set)
    }
}

fn scan_hwpx_security<R: Read + Seek>(
    file_names: &[String],
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
    doc: &mut Document,
) {
    let mut external_refs = Vec::new();
    let mut macro_modules = Vec::new();
    let mut has_autoexec = false;
    let mut encrypted_flag = false;
    let mut protected_flag = false;
    let mut diagnostics = Diagnostics::new();
    diagnostics.span = Some(SourceSpan::new("package"));
    for path in file_names {
        let lower = path.to_ascii_lowercase();
        if lower.starts_with("bindata/") {
            if lower.contains("ole") || lower.contains("object") {
                let mut ole = OleObject::new();
                ole.name = Some(path.clone());
                ole.size_bytes = zip.file_size(path).unwrap_or(0);
                store.insert(IRNode::OleObject(ole));
            }
        }

        if lower.contains("encrypt") {
            encrypted_flag = true;
        }
        if lower.contains("protect") || lower.contains("password") {
            protected_flag = true;
        }

        let is_script = lower.starts_with("scripts/")
            || lower.contains("/scripts/")
            || lower.starts_with("macros/")
            || lower.contains("/macros/")
            || lower.ends_with(".js")
            || lower.ends_with(".vbs")
            || lower.ends_with(".wsf")
            || lower.ends_with(".sct")
            || lower.ends_with(".py");
        if is_script {
            if let Ok(data) = zip.read_file(path) {
                let source = String::from_utf8_lossy(&data).to_string();
                let mut module = MacroModule::new(path.clone(), MacroModuleType::Standard);
                module.source_code = Some(source.clone());
                module.span = Some(SourceSpan::new(path));
                let id = module.id;
                store.insert(IRNode::MacroModule(module));
                macro_modules.push(id);
                let lower_source = source.to_ascii_lowercase();
                if lower.contains("auto")
                    || lower_source.contains("autoexec")
                    || lower_source.contains("auto_open")
                    || lower_source.contains("onopen")
                {
                    has_autoexec = true;
                }
            }
        }

        if !path.ends_with(".xml") {
            continue;
        }
        if let Ok(xml) = zip.read_file_string(path) {
            let refs = scan_hwpx_external_refs(&xml, path);
            external_refs.extend(refs);
            let lower_xml = xml.to_ascii_lowercase();
            if lower_xml.contains("encrypt") {
                encrypted_flag = true;
            }
            if lower_xml.contains("protect") || lower_xml.contains("password") {
                protected_flag = true;
            }
            if xml.to_ascii_lowercase().contains("ole") {
                let mut ole = OleObject::new();
                ole.name = Some(path.clone());
                ole.size_bytes = xml.len() as u64;
                store.insert(IRNode::OleObject(ole));
            }
        }
    }
    for ext in external_refs {
        store.insert(IRNode::ExternalReference(ext));
    }
    if !macro_modules.is_empty() {
        let mut project = MacroProject::new();
        project.name = Some("HWPX Scripts".to_string());
        project.modules = macro_modules;
        project.has_auto_exec = has_autoexec;
        project.span = Some(SourceSpan::new("package"));
        store.insert(IRNode::MacroProject(project));
    }
    if encrypted_flag {
        push_warning(
            &mut diagnostics,
            "HWPX_ENCRYPTED",
            "HWPX encrypted content detected".to_string(),
            None,
        );
    }
    if protected_flag {
        push_info(
            &mut diagnostics,
            "HWPX_PROTECTED",
            "HWPX protected content detected".to_string(),
            None,
        );
    }
    if !diagnostics.entries.is_empty() {
        let diag_id = diagnostics.id;
        store.insert(IRNode::Diagnostics(diagnostics));
        doc.diagnostics.push(diag_id);
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

const HWPTAG_BEGIN: u16 = 0x010;
const HWPTAG_DOCUMENT_PROPERTIES: u16 = HWPTAG_BEGIN;
const HWPTAG_PARA_HEADER: u16 = HWPTAG_BEGIN + 50;
const HWPTAG_PARA_TEXT: u16 = HWPTAG_BEGIN + 51;

struct HwpHeader {
    version: u32,
    flags: u32,
}

fn parse_file_header(data: &[u8]) -> Result<HwpHeader, ParseError> {
    if data.len() < 40 {
        return Err(ParseError::InvalidStructure(
            "HWP FileHeader too short".to_string(),
        ));
    }
    let signature = &data[..32];
    let signature = String::from_utf8_lossy(signature)
        .trim_matches('\0')
        .to_string();
    if !signature.contains("HWP Document File") {
        return Err(ParseError::InvalidStructure(format!(
            "Invalid HWP signature: {}",
            signature
        )));
    }
    let version = read_u32_le(data, 32)
        .ok_or_else(|| ParseError::InvalidStructure("Missing HWP version".to_string()))?;
    let flags = read_u32_le(data, 36)
        .ok_or_else(|| ParseError::InvalidStructure("Missing HWP flags".to_string()))?;
    Ok(HwpHeader { version, flags })
}

struct HwpRecord<'a> {
    tag_id: u16,
    level: u16,
    size: u32,
    data: &'a [u8],
}

fn for_each_record<F: FnMut(HwpRecord)>(data: &[u8], mut f: F) -> Result<(), ParseError> {
    let mut offset = 0usize;
    while offset + 4 <= data.len() {
        let header = read_u32_le(data, offset)
            .ok_or_else(|| ParseError::InvalidStructure("Invalid record header".to_string()))?;
        offset += 4;

        let tag_id = (header & 0x3FF) as u16;
        let level = ((header >> 10) & 0x3FF) as u16;
        let mut size = ((header >> 20) & 0xFFF) as u32;
        if size == 0xFFF {
            size = read_u32_le(data, offset).ok_or_else(|| {
                ParseError::InvalidStructure("Invalid extended record size".to_string())
            })?;
            offset += 4;
        }

        let end = offset + size as usize;
        if end > data.len() {
            return Err(ParseError::InvalidStructure(
                "Record size exceeds stream length".to_string(),
            ));
        }
        let payload = &data[offset..end];
        offset = end;
        f(HwpRecord {
            tag_id,
            level,
            size,
            data: payload,
        });
    }
    Ok(())
}

fn parse_docinfo_section_count(data: &[u8]) -> Result<Option<u16>, ParseError> {
    let mut section_count = None;
    for_each_record(data, |rec| {
        if rec.tag_id == HWPTAG_DOCUMENT_PROPERTIES && rec.data.len() >= 2 {
            section_count = read_u16_le(rec.data, 0);
        }
    })?;
    Ok(section_count)
}

fn parse_hwp_section_stream(
    data: &[u8],
    source: &str,
    store: &mut IrStore,
) -> Result<Vec<NodeId>, ParseError> {
    let mut paragraphs: Vec<NodeId> = Vec::new();
    let mut current_para: Option<Paragraph> = None;

    for_each_record(data, |rec| match rec.tag_id {
        HWPTAG_PARA_HEADER => {
            if let Some(para) = current_para.take() {
                let para_id = para.id;
                store.insert(IRNode::Paragraph(para));
                paragraphs.push(para_id);
            }
            let mut para = Paragraph::new();
            para.span = Some(SourceSpan::new(source));
            let _text_count = parse_para_text_count(rec.data);
            current_para = Some(para);
        }
        HWPTAG_PARA_TEXT => {
            let text = decode_hwp_text(rec.data);
            if !text.is_empty() {
                let mut run = Run::new(text);
                run.span = Some(SourceSpan::new(source));
                let run_id = run.id;
                store.insert(IRNode::Run(run));
                if let Some(para) = current_para.as_mut() {
                    para.runs.push(run_id);
                } else {
                    let mut para = Paragraph::new();
                    para.span = Some(SourceSpan::new(source));
                    para.runs.push(run_id);
                    current_para = Some(para);
                }
            }
        }
        _ => {}
    })?;

    if let Some(para) = current_para.take() {
        let para_id = para.id;
        store.insert(IRNode::Paragraph(para));
        paragraphs.push(para_id);
    }

    Ok(paragraphs)
}

fn parse_para_text_count(data: &[u8]) -> Option<u32> {
    let mut count = read_u32_le(data, 0)?;
    if count & 0x8000_0000 != 0 {
        count &= 0x7FFF_FFFF;
    }
    Some(count)
}

fn decode_hwp_text(data: &[u8]) -> String {
    if data.is_empty() {
        return String::new();
    }
    let text = if data.len() % 2 == 0 {
        decode_utf16le(data)
    } else {
        String::from_utf8_lossy(data).to_string()
    };
    sanitize_text(&text)
}

fn sanitize_text(value: &str) -> String {
    value
        .chars()
        .map(|c| if c <= '\u{001F}' { ' ' } else { c })
        .collect::<String>()
        .trim()
        .to_string()
}

fn decode_utf16le(data: &[u8]) -> String {
    let mut units = Vec::with_capacity(data.len() / 2);
    for chunk in data.chunks(2) {
        if chunk.len() == 2 {
            units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
    }
    String::from_utf16_lossy(&units)
}

fn maybe_decompress_stream(
    data: &[u8],
    compressed: bool,
    source: &str,
) -> Result<Vec<u8>, ParseError> {
    if !compressed {
        return Ok(data.to_vec());
    }
    if data.len() < 2 {
        return Ok(data.to_vec());
    }
    let mut out = Vec::new();
    let zlib_result = {
        let mut decoder = ZlibDecoder::new(data);
        decoder.read_to_end(&mut out)
    };
    if zlib_result.is_ok() {
        return Ok(out);
    }
    out.clear();
    let mut decoder = DeflateDecoder::new(data);
    decoder.read_to_end(&mut out).map_err(|e| {
        ParseError::InvalidStructure(format!("Failed to decompress HWP stream {}: {e}", source))
    })?;
    Ok(out)
}

fn derive_hwp_key(password: &str) -> [u8; 16] {
    let mut hasher = Sha1::new();
    hasher.update(password.as_bytes());
    let digest = hasher.finalize();
    let mut key = [0u8; 16];
    key.copy_from_slice(&digest[..16]);
    key
}

fn decrypt_hwp_stream(data: &[u8], password: &str, source: &str) -> Result<Vec<u8>, ParseError> {
    if data.len() % 16 != 0 {
        return Err(ParseError::InvalidStructure(format!(
            "HWP encrypted stream {} length not aligned",
            source
        )));
    }
    let key = derive_hwp_key(password);
    let iv = [0u8; 16];
    let mut buffer = data.to_vec();
    let decryptor = Decryptor::<Aes128>::new_from_slices(&key, &iv).map_err(|e| {
        ParseError::InvalidStructure(format!(
            "Failed to init HWP decryptor for {}: {}",
            source, e
        ))
    })?;
    let decrypted = decryptor
        .decrypt_padded_mut::<Pkcs7>(&mut buffer)
        .map_err(|e| {
            ParseError::InvalidStructure(format!("Failed to decrypt HWP stream {}: {}", source, e))
        })?;
    Ok(decrypted.to_vec())
}

fn prepare_hwp_stream_data(
    data: &[u8],
    encrypted: bool,
    password: Option<&str>,
    force_parse: bool,
    try_raw_encrypted: bool,
    source: &str,
    diagnostics: &mut Diagnostics,
) -> Option<Vec<u8>> {
    if !encrypted {
        return Some(data.to_vec());
    }
    if let Some(password) = password {
        match decrypt_hwp_stream(data, password, source) {
            Ok(bytes) => return Some(bytes),
            Err(err) => {
                push_warning(
                    diagnostics,
                    "HWP_DECRYPT_FAIL",
                    err.to_string(),
                    Some(source),
                );
            }
        }
    }
    if force_parse {
        push_warning(
            diagnostics,
            "HWP_FORCE_PARSE_STREAM",
            "HWP force-parse: using raw encrypted stream bytes".to_string(),
            Some(source),
        );
        return Some(data.to_vec());
    }
    if try_raw_encrypted {
        push_warning(
            diagnostics,
            "HWP_ENCRYPTED_RAW_STREAM",
            "HWP encrypted without password: trying raw stream bytes".to_string(),
            Some(source),
        );
        return Some(data.to_vec());
    }
    None
}

fn dump_hwp_streams(
    cfb: &Cfb,
    stream_names: &[String],
    compressed: bool,
    encrypted: bool,
    password: Option<&str>,
    force_parse: bool,
    try_raw_encrypted: bool,
    diagnostics: &mut Diagnostics,
) {
    for path in stream_names {
        let size = cfb.stream_size(path).unwrap_or(0);
        let mut sha = Sha256::new();
        let mut hash_hex = "missing".to_string();
        let mut decompress_status = "skip".to_string();
        if let Some(data) = cfb.read_stream(path) {
            sha.update(&data);
            let hash = sha.finalize();
            let mut out = String::with_capacity(hash.len() * 2);
            for byte in hash {
                out.push_str(&format!("{:02x}", byte));
            }
            hash_hex = out;

            if compressed {
                if let Some(bytes) = prepare_hwp_stream_data(
                    &data,
                    encrypted,
                    password,
                    force_parse,
                    try_raw_encrypted,
                    path,
                    diagnostics,
                ) {
                    match maybe_decompress_stream(&bytes, compressed, path) {
                        Ok(_) => decompress_status = "ok".to_string(),
                        Err(err) => {
                            decompress_status = format!("fail: {}", err);
                        }
                    }
                } else {
                    decompress_status = "encrypted".to_string();
                }
            }
        }
        push_info(
            diagnostics,
            "HWP_STREAM_DUMP",
            format!(
                "stream: {}, size={}, compressed={}, sha256={}, decompress={}",
                path, size, compressed, hash_hex, decompress_status
            ),
            Some(path),
        );
    }
}
fn parse_default_jscript(data: &[u8], store: &mut IrStore, source: &str) -> Option<NodeId> {
    let (name, source_code) = parse_jscript_stream(data)?;
    let mut module = MacroModule::new(name, MacroModuleType::Standard);
    module.source_code = Some(source_code);
    module.span = Some(SourceSpan::new(source));
    let module_id = module.id;
    store.insert(IRNode::MacroModule(module));

    let mut project = MacroProject::new();
    project.name = Some("HWP Script".to_string());
    project.modules.push(module_id);
    project.span = Some(SourceSpan::new(source));
    let project_id = project.id;
    store.insert(IRNode::MacroProject(project));

    Some(project_id)
}

fn parse_jscript_stream(data: &[u8]) -> Option<(String, String)> {
    if data.len() < 8 {
        return None;
    }
    let mut offset = 4;
    let mut strings = Vec::new();
    for _ in 0..3 {
        let (value, used) = read_len_string(data, offset)?;
        offset += used;
        if !value.is_empty() {
            strings.push(value);
        }
    }
    let name = strings
        .get(0)
        .cloned()
        .unwrap_or_else(|| "DefaultJScript".to_string());
    let source = strings.last().cloned().unwrap_or_else(|| String::new());
    if source.is_empty() {
        None
    } else {
        Some((name, source))
    }
}

fn extract_urls(text: &str) -> Vec<String> {
    let mut out = Vec::new();
    let patterns: [&[u8]; 5] = [b"http://", b"https://", b"file://", b"ftp://", b"mailto:"];
    let bytes = text.as_bytes();
    let mut idx = 0usize;
    while idx < bytes.len() {
        let mut matched = false;
        for pat in &patterns {
            if bytes[idx..].starts_with(pat) {
                matched = true;
                break;
            }
        }
        if matched {
            let mut end = idx;
            while end < bytes.len() {
                let ch = bytes[end];
                if ch.is_ascii_whitespace() || ch == b'"' || ch == b'\'' || ch == b')' || ch == b'>'
                {
                    break;
                }
                end += 1;
            }
            if end > idx {
                let url = String::from_utf8_lossy(&bytes[idx..end]).to_string();
                out.push(url);
            }
            idx = end;
        } else {
            idx += 1;
        }
    }
    out
}

fn scan_hwp_external_refs(
    cfb: &Cfb,
    stream_names: &[String],
    compressed: bool,
    encrypted: bool,
    password: Option<&str>,
    force_parse: bool,
    try_raw_encrypted: bool,
    diagnostics: &mut Diagnostics,
) -> Vec<ExternalReference> {
    let mut refs = Vec::new();
    for path in stream_names {
        if !path.starts_with("BodyText/Section") {
            continue;
        }
        if let Some(data) = cfb.read_stream(path) {
            let data = match prepare_hwp_stream_data(
                &data,
                encrypted,
                password,
                force_parse,
                try_raw_encrypted,
                path,
                diagnostics,
            ) {
                Some(bytes) => bytes,
                None => continue,
            };
            let data = match maybe_decompress_stream(&data, compressed, path) {
                Ok(bytes) => bytes,
                Err(err) => {
                    push_warning(
                        diagnostics,
                        "HWP_DECOMPRESS_FAIL",
                        err.to_string(),
                        Some(path),
                    );
                    continue;
                }
            };
            let text = decode_hwp_text(&data);
            for url in extract_urls(&text) {
                let mut ext = ExternalReference::new(ExternalRefType::Hyperlink, url);
                ext.span = Some(SourceSpan::new(path));
                refs.push(ext);
            }
        }
    }
    refs
}

fn read_len_string(data: &[u8], offset: usize) -> Option<(String, usize)> {
    let len = read_u32_le(data, offset)? as usize;
    let start = offset + 4;
    let mut bytes_len = len;
    if start + bytes_len > data.len() {
        let alt = len * 2;
        if start + alt <= data.len() {
            bytes_len = alt;
        } else {
            return None;
        }
    }
    let bytes = &data[start..start + bytes_len];
    let text = decode_string_bytes(bytes);
    Some((text, 4 + bytes_len))
}

fn decode_string_bytes(bytes: &[u8]) -> String {
    let zero_bytes = bytes.iter().filter(|b| **b == 0).count();
    if bytes.len() % 2 == 0 && zero_bytes > 0 {
        decode_utf16le(bytes)
    } else {
        String::from_utf8_lossy(bytes).to_string()
    }
}

fn read_u16_le(data: &[u8], offset: usize) -> Option<u16> {
    if offset + 2 > data.len() {
        return None;
    }
    Some(u16::from_le_bytes([data[offset], data[offset + 1]]))
}

fn read_u32_le(data: &[u8], offset: usize) -> Option<u32> {
    if offset + 4 > data.len() {
        return None;
    }
    Some(u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ]))
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
