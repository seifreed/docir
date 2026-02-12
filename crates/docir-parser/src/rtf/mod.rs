//! RTF parsing support.

use crate::error::ParseError;
use crate::format::FormatParser;
use crate::input::{parse_from_bytes, parse_from_file, read_all_with_limit};
use crate::parser::{ParsedDocument, ParserConfig};
use docir_core::ir::ParagraphBorders;
use docir_core::ir::{
    Border, BorderStyle, CellVerticalAlignment, Indentation, LineSpacingRule, MergeType, Spacing,
    TextAlignment,
};
use docir_core::ir::{
    Field, FieldInstruction, FieldKind, Hyperlink, IRNode, MediaType, NumberingInfo, Paragraph,
    Run, RunProperties, Style, StyleSet, StyleType, Table, TableCell, TableCellProperties,
    TableRow, TableWidth, TableWidthType,
};
use docir_core::normalize::normalize_store;
use docir_core::security::{ExternalRefType, ExternalReference};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use encoding_rs::Encoding;
use std::collections::HashMap;
use std::io::{Read, Seek};
use std::path::Path;

mod objects;

use self::objects::{finalize_object, finalize_picture, ObjectContext, ObjectTextTarget};

/// Parser for RTF documents.
pub struct RtfParser {
    config: ParserConfig,
}

impl FormatParser for RtfParser {
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        self.parse_reader(reader)
    }
}

impl RtfParser {
    /// Creates a new parser with default configuration.
    pub fn new() -> Self {
        Self {
            config: ParserConfig::default(),
        }
    }

    /// Creates a new parser with custom configuration.
    pub fn with_config(config: ParserConfig) -> Self {
        Self { config }
    }

    /// Parses a file from the filesystem.
    pub fn parse_file<P: AsRef<Path>>(&self, path: P) -> Result<ParsedDocument, ParseError> {
        parse_from_file(path, |reader| self.parse_reader(reader))
    }

    /// Parses from a byte slice.
    pub fn parse_bytes(&self, data: &[u8]) -> Result<ParsedDocument, ParseError> {
        parse_from_bytes(data, |reader| self.parse_reader(reader))
    }

    /// Parses from any reader.
    pub fn parse_reader<R: Read + Seek>(
        &self,
        mut reader: R,
    ) -> Result<ParsedDocument, ParseError> {
        let data = read_all_with_limit(reader, self.config.max_input_size)?;
        if !is_rtf_bytes(&data) {
            return Err(ParseError::UnsupportedFormat(
                "Missing RTF header".to_string(),
            ));
        }

        let mut store = IrStore::new();
        let mut doc = docir_core::ir::Document::new(DocumentFormat::Rtf);

        let mut ctx = RtfParseContext::new(
            self.config.rtf_max_group_depth,
            self.config.rtf_max_object_hex_len,
        );
        let mut cursor = RtfCursor::new(&data);
        parse_rtf(&mut cursor, &mut ctx, &mut store)?;

        if let Some(style_set) = ctx.style_set.take() {
            let style_id = style_set.id;
            store.insert(IRNode::StyleSet(style_set));
            doc.styles = Some(style_id);
        }

        for section in ctx.sections {
            doc.content.push(section);
        }
        for media in ctx.media_assets {
            doc.shared_parts.push(media);
        }
        for ext in ctx.external_refs {
            doc.security.external_refs.push(ext);
        }
        for ole in ctx.ole_objects {
            doc.security.ole_objects.push(ole);
        }

        let root_id = doc.id;
        store.insert(IRNode::Document(doc));
        normalize_store(&mut store, root_id);

        Ok(ParsedDocument {
            root_id,
            format: DocumentFormat::Rtf,
            store,
            metrics: None,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum GroupKind {
    Normal,
    FontTable,
    ColorTable,
    Stylesheet,
    StylesheetEntry,
    Info,
    Skip,
    Field,
    FieldInst,
    FieldResult,
    Picture,
    Object,
    ListTable,
    List,
    ListLevel,
    ListOverrideTable,
    ListOverride,
}

#[derive(Debug, Clone, Default)]
struct RtfStyleState {
    style_id: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strike: Option<bool>,
    font_size: Option<u32>,
    font_index: Option<u32>,
    color_index: Option<usize>,
    highlight_index: Option<usize>,
    vertical: Option<docir_core::ir::VerticalTextAlignment>,
}

#[derive(Debug, Clone)]
struct GroupState {
    kind: GroupKind,
    style: RtfStyleState,
}

#[derive(Debug, Default)]
struct FontTable {
    fonts: HashMap<u32, String>,
}

#[derive(Debug, Default)]
struct ColorTable {
    colors: Vec<Option<String>>, // index 0 is default
}

#[derive(Debug, Default)]
struct FieldContext {
    instruction: String,
    runs: Vec<NodeId>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BorderTarget {
    Top,
    Bottom,
    Left,
    Right,
    InsideH,
    InsideV,
}

#[derive(Debug, Default)]
struct StyleEntryContext {
    style_id: Option<String>,
    style_type: Option<StyleType>,
    name_buf: String,
}

#[derive(Debug)]
struct RtfParseContext {
    sections: Vec<NodeId>,
    current_section: Option<NodeId>,
    current_paragraph: Option<Paragraph>,
    current_table: Option<Table>,
    current_row: Option<TableRow>,
    current_cell: Option<TableCell>,
    font_table: FontTable,
    color_table: ColorTable,
    encoding: &'static Encoding,
    group_stack: Vec<GroupState>,
    field_stack: Vec<FieldContext>,
    object_stack: Vec<ObjectContext>,
    style_set: Option<StyleSet>,
    current_style: Option<StyleEntryContext>,
    list_overrides: HashMap<i32, i32>,
    list_levels: HashMap<i32, u32>,
    list_level_formats: HashMap<(i32, u32), String>,
    current_list_id: Option<i32>,
    current_list_level: u32,
    current_list_override: Option<i32>,
    current_list_override_list_id: Option<i32>,
    pending_list_override: Option<i32>,
    pending_list_level: Option<u32>,
    pending_para_style: Option<String>,
    pending_alignment: Option<TextAlignment>,
    pending_indent: Indentation,
    pending_spacing: Spacing,
    pending_line_rule: Option<LineSpacingRule>,
    pending_para_border_target: Option<BorderTarget>,
    pending_para_border: Border,
    pending_para_borders: ParagraphBorders,
    row_cellx: Vec<i32>,
    current_cell_index: usize,
    pending_cell_props: Option<TableCellProperties>,
    pending_border_target: Option<BorderTarget>,
    pending_border: Border,
    object_text_target: Option<ObjectTextTarget>,
    media_assets: Vec<NodeId>,
    external_refs: Vec<NodeId>,
    ole_objects: Vec<NodeId>,
    current_text: String,
    current_props: RtfStyleState,
    max_group_depth: usize,
    max_object_hex_len: usize,
}

impl RtfParseContext {
    fn new(max_group_depth: usize, max_object_hex_len: usize) -> Self {
        let mut ctx = Self {
            sections: Vec::new(),
            current_section: None,
            current_paragraph: None,
            current_table: None,
            current_row: None,
            current_cell: None,
            font_table: FontTable::default(),
            color_table: ColorTable::default(),
            encoding: encoding_rs::WINDOWS_1252,
            group_stack: Vec::new(),
            field_stack: Vec::new(),
            object_stack: Vec::new(),
            style_set: None,
            current_style: None,
            list_overrides: HashMap::new(),
            list_levels: HashMap::new(),
            list_level_formats: HashMap::new(),
            current_list_id: None,
            current_list_level: 0,
            current_list_override: None,
            current_list_override_list_id: None,
            pending_list_override: None,
            pending_list_level: None,
            pending_para_style: None,
            pending_alignment: None,
            pending_indent: Indentation::default(),
            pending_spacing: Spacing::default(),
            pending_line_rule: None,
            pending_para_border_target: None,
            pending_para_border: Border {
                style: BorderStyle::None,
                width: None,
                color: None,
            },
            pending_para_borders: ParagraphBorders::default(),
            row_cellx: Vec::new(),
            current_cell_index: 0,
            pending_cell_props: None,
            pending_border_target: None,
            pending_border: Border {
                style: BorderStyle::None,
                width: None,
                color: None,
            },
            object_text_target: None,
            media_assets: Vec::new(),
            external_refs: Vec::new(),
            ole_objects: Vec::new(),
            current_text: String::new(),
            current_props: RtfStyleState::default(),
            max_group_depth,
            max_object_hex_len,
        };
        ctx.group_stack.push(GroupState {
            kind: GroupKind::Normal,
            style: RtfStyleState::default(),
        });
        ctx
    }

    fn current_group_kind(&self) -> GroupKind {
        self.group_stack
            .last()
            .map(|g| g.kind)
            .unwrap_or(GroupKind::Normal)
    }

    fn push_group(&mut self, kind: GroupKind) -> Result<(), ParseError> {
        if self.max_group_depth > 0 && self.group_stack.len() >= self.max_group_depth {
            return Err(ParseError::ResourceLimit(format!(
                "RTF max group depth exceeded: {}",
                self.max_group_depth
            )));
        }
        let style = self.current_props.clone();
        self.group_stack.push(GroupState { kind, style });
        Ok(())
    }

    fn pop_group(&mut self) {
        if let Some(group) = self.group_stack.pop() {
            self.current_props = group.style;
        }
    }
}

struct RtfCursor<'a> {
    data: &'a [u8],
    pos: usize,
}

impl<'a> RtfCursor<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, pos: 0 }
    }

    fn peek(&self) -> Option<u8> {
        self.data.get(self.pos).copied()
    }

    fn next(&mut self) -> Option<u8> {
        let b = self.data.get(self.pos).copied();
        if b.is_some() {
            self.pos += 1;
        }
        b
    }

    fn is_eof(&self) -> bool {
        self.pos >= self.data.len()
    }
}

pub(crate) fn is_rtf_bytes(data: &[u8]) -> bool {
    data.starts_with(b"{\\rtf") || data.starts_with(b"{\rtf")
}

fn parse_rtf(
    cursor: &mut RtfCursor<'_>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<(), ParseError> {
    while !cursor.is_eof() {
        match cursor.next() {
            Some(b'{') => {
                ctx.push_group(GroupKind::Normal)?;
            }
            Some(b'}') => {
                flush_text(ctx, store, None)?;
                match ctx.current_group_kind() {
                    GroupKind::Field => finalize_field(ctx, store),
                    GroupKind::Picture => {
                        if let Some(obj) = ctx.object_stack.pop() {
                            if let Some(asset_id) = finalize_picture(obj, store) {
                                ctx.media_assets.push(asset_id);
                            }
                        }
                    }
                    GroupKind::Object => {
                        if let Some(obj) = ctx.object_stack.pop() {
                            if let Some(ole_id) = finalize_object(obj, store) {
                                ctx.ole_objects.push(ole_id);
                            }
                        }
                    }
                    GroupKind::StylesheetEntry => {
                        finalize_style_entry(ctx);
                    }
                    GroupKind::List => {
                        finalize_list_entry(ctx);
                    }
                    GroupKind::ListOverride => {
                        finalize_list_override(ctx);
                    }
                    _ => {}
                }
                ctx.pop_group();
            }
            Some(b'\\') => {
                parse_control(cursor, ctx, store)?;
            }
            Some(b'\r') | Some(b'\n') => {
                // ignore raw newlines
            }
            Some(byte) => match ctx.current_group_kind() {
                GroupKind::Normal | GroupKind::FieldResult | GroupKind::FieldInst => {
                    append_text_byte(ctx, byte);
                }
                GroupKind::Object | GroupKind::Picture => {
                    if byte.is_ascii_hexdigit() {
                        if let Some(obj) = ctx.object_stack.last_mut() {
                            obj.data_hex_len += 1;
                            if ctx.max_object_hex_len > 0
                                && obj.data_hex_len > ctx.max_object_hex_len
                            {
                                return Err(ParseError::ResourceLimit(format!(
                                    "RTF objdata too large: {} hex chars (max: {})",
                                    obj.data_hex_len, ctx.max_object_hex_len
                                )));
                            }
                        }
                    }
                }
                GroupKind::FontTable | GroupKind::ColorTable => {
                    append_text_byte(ctx, byte);
                }
                _ => {}
            },
            None => break,
        }
    }
    flush_text(ctx, store, None)?;
    finalize_table_if_open(ctx, store);
    finalize_paragraph(ctx, store);
    finalize_section(ctx, store);
    Ok(())
}

fn parse_control(
    cursor: &mut RtfCursor<'_>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<(), ParseError> {
    let Some(next) = cursor.next() else {
        return Ok(());
    };

    match next {
        b'\\' | b'{' | b'}' => {
            append_text_byte(ctx, next);
            return Ok(());
        }
        b'\'' => {
            let hi = cursor.next().unwrap_or(b'0');
            let lo = cursor.next().unwrap_or(b'0');
            if let (Some(h), Some(l)) = (hex_val(hi), hex_val(lo)) {
                append_text_byte(ctx, (h << 4) | l);
            }
            return Ok(());
        }
        b'*' => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Skip;
            }
            return Ok(());
        }
        b'~' => {
            append_text(ctx, " ");
            return Ok(());
        }
        b'-' => {
            return Ok(()); // optional hyphen
        }
        b'_' => {
            append_text(ctx, "-");
            return Ok(());
        }
        b'\n' | b'\r' => return Ok(()),
        _ => {}
    }

    if next.is_ascii_alphabetic() {
        let mut word = Vec::new();
        word.push(next);
        while let Some(b) = cursor.peek() {
            if b.is_ascii_alphabetic() {
                word.push(b);
                cursor.next();
            } else {
                break;
            }
        }

        let mut sign = 1i32;
        let mut param: Option<i32> = None;
        if let Some(b) = cursor.peek() {
            if b == b'-' {
                sign = -1;
                cursor.next();
            }
        }
        let mut digits = Vec::new();
        while let Some(b) = cursor.peek() {
            if b.is_ascii_digit() {
                digits.push(b);
                cursor.next();
            } else {
                break;
            }
        }
        if !digits.is_empty() {
            if let Ok(num) = std::str::from_utf8(&digits).unwrap_or("0").parse::<i32>() {
                param = Some(num * sign);
            }
        }

        if let Some(b) = cursor.peek() {
            if b == b' ' {
                cursor.next();
            }
        }

        let word = String::from_utf8_lossy(&word).to_ascii_lowercase();
        handle_control_word(&word, param, ctx, store)?;
    }

    Ok(())
}

fn handle_control_word(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<(), ParseError> {
    if handle_paragraph_controls(word, param, ctx, store)? {
        return Ok(());
    }
    if handle_run_style_controls(word, param, ctx) {
        return Ok(());
    }
    if handle_group_controls(word, param, ctx)? {
        return Ok(());
    }
    if handle_field_controls(word, ctx) {
        return Ok(());
    }
    if handle_object_controls(word, param, ctx) {
        return Ok(());
    }
    if handle_table_controls(word, param, ctx, store)? {
        return Ok(());
    }
    if handle_encoding_controls(word, param, ctx) {
        return Ok(());
    }
    Ok(())
}

fn handle_group_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
) -> Result<bool, ParseError> {
    match word {
        "fonttbl" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::FontTable;
            }
            ctx.current_text.clear();
        }
        "colortbl" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::ColorTable;
            }
            if ctx.color_table.colors.is_empty() {
                ctx.color_table.colors.push(None);
            }
            ctx.current_text.clear();
        }
        "stylesheet" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Stylesheet;
            }
            if ctx.style_set.is_none() {
                ctx.style_set = Some(StyleSet::new());
            }
        }
        "s" => {
            if let Some(id) = param {
                let style_id = format!("s{}", id.max(0));
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    ctx.group_stack
                        .last_mut()
                        .map(|g| g.kind = GroupKind::StylesheetEntry);
                    ctx.current_style = Some(StyleEntryContext {
                        style_id: Some(style_id),
                        style_type: Some(StyleType::Paragraph),
                        name_buf: String::new(),
                    });
                } else {
                    ctx.pending_para_style = Some(style_id);
                    if let Some(para) = ctx.current_paragraph.as_mut() {
                        para.style_id = ctx.pending_para_style.clone();
                    }
                }
            }
        }
        "cs" => {
            if let Some(id) = param {
                let style_id = format!("cs{}", id.max(0));
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    ctx.group_stack
                        .last_mut()
                        .map(|g| g.kind = GroupKind::StylesheetEntry);
                    ctx.current_style = Some(StyleEntryContext {
                        style_id: Some(style_id),
                        style_type: Some(StyleType::Character),
                        name_buf: String::new(),
                    });
                } else {
                    ctx.current_props.style_id = Some(style_id);
                }
            }
        }
        "ds" => {
            if let Some(id) = param {
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    ctx.group_stack
                        .last_mut()
                        .map(|g| g.kind = GroupKind::StylesheetEntry);
                    ctx.current_style = Some(StyleEntryContext {
                        style_id: Some(format!("ds{}", id.max(0))),
                        style_type: Some(StyleType::Other),
                        name_buf: String::new(),
                    });
                }
            }
        }
        "ts" => {
            if let Some(id) = param {
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    ctx.group_stack
                        .last_mut()
                        .map(|g| g.kind = GroupKind::StylesheetEntry);
                    ctx.current_style = Some(StyleEntryContext {
                        style_id: Some(format!("ts{}", id.max(0))),
                        style_type: Some(StyleType::Table),
                        name_buf: String::new(),
                    });
                }
            }
        }
        "info" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Info;
            }
        }
        "listtable" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::ListTable;
            }
        }
        "listoverridetable" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::ListOverrideTable;
            }
        }
        "list" => {
            if ctx.current_group_kind() == GroupKind::ListTable {
                if let Some(group) = ctx.group_stack.last_mut() {
                    group.kind = GroupKind::List;
                }
                ctx.current_list_id = None;
                ctx.current_list_level = 0;
            }
        }
        "listlevel" => {
            if ctx.current_group_kind() == GroupKind::List {
                ctx.current_list_level = ctx.current_list_level.saturating_add(1);
                if let Some(group) = ctx.group_stack.last_mut() {
                    group.kind = GroupKind::ListLevel;
                }
            }
        }
        "levelnfc" => {
            if let (Some(list_id), Some(nfc)) = (ctx.current_list_id, param) {
                let level = ctx.current_list_level.saturating_sub(1);
                ctx.list_level_formats
                    .insert((list_id, level), format!("nfc:{}", nfc));
            }
        }
        "listid" => {
            if let Some(id) = param {
                if matches!(
                    ctx.current_group_kind(),
                    GroupKind::List | GroupKind::ListLevel | GroupKind::ListOverride
                ) {
                    ctx.current_list_id = Some(id);
                    if ctx.current_group_kind() == GroupKind::ListOverride {
                        ctx.current_list_override_list_id = Some(id);
                    }
                }
            }
        }
        "listoverride" => {
            if ctx.current_group_kind() == GroupKind::ListOverrideTable {
                if let Some(group) = ctx.group_stack.last_mut() {
                    group.kind = GroupKind::ListOverride;
                }
                ctx.current_list_override = None;
                ctx.current_list_override_list_id = None;
            }
        }
        "ls" => {
            if let Some(id) = param {
                if ctx.current_group_kind() == GroupKind::ListOverride {
                    ctx.current_list_override = Some(id);
                } else {
                    ctx.pending_list_override = Some(id);
                    let numbering = pending_numbering(ctx);
                    if let Some(para) = ctx.current_paragraph.as_mut() {
                        if let Some(numbering) = numbering {
                            para.properties.numbering = Some(numbering);
                        }
                    }
                }
            }
        }
        "ilvl" => {
            if let Some(level) = param {
                ctx.pending_list_level = Some(level.max(0) as u32);
                let numbering = pending_numbering(ctx);
                if let Some(para) = ctx.current_paragraph.as_mut() {
                    if let Some(numbering) = numbering {
                        para.properties.numbering = Some(numbering);
                    }
                }
            }
        }
        _ => return Ok(false),
    }
    Ok(true)
}

fn handle_field_controls(word: &str, ctx: &mut RtfParseContext) -> bool {
    match word {
        "field" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Field;
            }
            ctx.field_stack.push(FieldContext::default());
        }
        "fldinst" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::FieldInst;
            }
            ctx.current_text.clear();
        }
        "fldrslt" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::FieldResult;
            }
            ctx.current_text.clear();
        }
        _ => return false,
    }
    true
}

fn handle_object_controls(word: &str, param: Option<i32>, ctx: &mut RtfParseContext) -> bool {
    match word {
        "pict" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Picture;
            }
            ctx.object_stack.push(ObjectContext::default());
        }
        "pngblip" => {
            if let Some(obj) = ctx.object_stack.last_mut() {
                obj.media_type = Some(MediaType::Image);
            }
        }
        "jpegblip" | "jpgblip" => {
            if let Some(obj) = ctx.object_stack.last_mut() {
                obj.media_type = Some(MediaType::Image);
            }
        }
        "wmetafile" | "emfblip" | "wmetafile8" => {
            if let Some(obj) = ctx.object_stack.last_mut() {
                obj.media_type = Some(MediaType::Image);
            }
        }
        "picw" => {
            if let Some(value) = param {
                if let Some(obj) = ctx.object_stack.last_mut() {
                    obj.pic_width = Some(value.max(0) as u32);
                }
            }
        }
        "pich" => {
            if let Some(value) = param {
                if let Some(obj) = ctx.object_stack.last_mut() {
                    obj.pic_height = Some(value.max(0) as u32);
                }
            }
        }
        "picwgoal" => {
            if let Some(value) = param {
                if let Some(obj) = ctx.object_stack.last_mut() {
                    obj.pic_width = Some(value.max(0) as u32);
                }
            }
        }
        "pichgoal" => {
            if let Some(value) = param {
                if let Some(obj) = ctx.object_stack.last_mut() {
                    obj.pic_height = Some(value.max(0) as u32);
                }
            }
        }
        "object" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Object;
            }
            ctx.object_stack.push(ObjectContext::default());
        }
        "objclass" => {
            ctx.current_text.clear();
            ctx.object_text_target = Some(ObjectTextTarget::Class);
        }
        "objname" => {
            ctx.current_text.clear();
            ctx.object_text_target = Some(ObjectTextTarget::Name);
        }
        "objdata" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Object;
            }
            ctx.current_text.clear();
            ctx.object_text_target = None;
        }
        _ => return false,
    }
    true
}

fn handle_table_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<bool, ParseError> {
    match word {
        "trowd" => {
            flush_text(ctx, store, None)?;
            ensure_section(ctx, store);
            if ctx.current_table.is_none() {
                ctx.current_table = Some(Table::new());
            }
            ctx.current_row = Some(TableRow::new());
            ctx.current_cell = Some(TableCell::new());
            ctx.row_cellx.clear();
            ctx.current_cell_index = 0;
            ctx.pending_cell_props = None;
            ctx.pending_border_target = None;
        }
        "cellx" => {
            if let Some(value) = param {
                ctx.row_cellx.push(value);
            }
        }
        "clvmgf" => {
            ctx.pending_cell_props
                .get_or_insert_with(TableCellProperties::default)
                .vertical_merge = Some(MergeType::Restart);
        }
        "clvmrg" => {
            ctx.pending_cell_props
                .get_or_insert_with(TableCellProperties::default)
                .vertical_merge = Some(MergeType::Continue);
        }
        "clgridspan" => {
            if let Some(value) = param {
                ctx.pending_cell_props
                    .get_or_insert_with(TableCellProperties::default)
                    .grid_span = Some(value.max(1) as u32);
            }
        }
        "clcbpat" => {
            if let Some(index) = param {
                if let Some(color) = color_from_index(ctx, index as usize) {
                    ctx.pending_cell_props
                        .get_or_insert_with(TableCellProperties::default)
                        .shading = Some(color);
                }
            }
        }
        "clvertalt" => {
            ctx.pending_cell_props
                .get_or_insert_with(TableCellProperties::default)
                .vertical_align = Some(CellVerticalAlignment::Top);
        }
        "clvertalc" => {
            ctx.pending_cell_props
                .get_or_insert_with(TableCellProperties::default)
                .vertical_align = Some(CellVerticalAlignment::Center);
        }
        "clvertalb" => {
            ctx.pending_cell_props
                .get_or_insert_with(TableCellProperties::default)
                .vertical_align = Some(CellVerticalAlignment::Bottom);
        }
        "clbrdrt" => set_border_target(ctx, BorderTarget::Top),
        "clbrdrb" => set_border_target(ctx, BorderTarget::Bottom),
        "clbrdrl" => set_border_target(ctx, BorderTarget::Left),
        "clbrdrr" => set_border_target(ctx, BorderTarget::Right),
        "clbrdrh" => set_border_target(ctx, BorderTarget::InsideH),
        "clbrdrv" => set_border_target(ctx, BorderTarget::InsideV),
        "brdrs" => {
            ctx.pending_border.style = BorderStyle::Single;
            apply_border(ctx);
        }
        "brdrth" => {
            ctx.pending_border.style = BorderStyle::Thick;
            apply_border(ctx);
        }
        "brdrdb" => {
            ctx.pending_border.style = BorderStyle::Double;
            apply_border(ctx);
        }
        "brdrdot" => {
            ctx.pending_border.style = BorderStyle::Dotted;
            apply_border(ctx);
        }
        "brdrdash" => {
            ctx.pending_border.style = BorderStyle::Dashed;
            apply_border(ctx);
        }
        "brdrtriple" => {
            ctx.pending_border.style = BorderStyle::Triple;
            apply_border(ctx);
        }
        "brdrw" => {
            if let Some(value) = param {
                ctx.pending_border.width = Some(value.max(0) as u32);
                apply_border(ctx);
            }
        }
        "brdrcf" => {
            if let Some(index) = param {
                if let Some(color) = color_from_index(ctx, index as usize) {
                    ctx.pending_border.color = Some(color);
                    apply_border(ctx);
                }
            }
        }
        "brdrt" => {
            ctx.pending_para_border_target = Some(BorderTarget::Top);
            apply_paragraph_border(ctx);
        }
        "brdrb" => {
            ctx.pending_para_border_target = Some(BorderTarget::Bottom);
            apply_paragraph_border(ctx);
        }
        "brdrl" => {
            ctx.pending_para_border_target = Some(BorderTarget::Left);
            apply_paragraph_border(ctx);
        }
        "brdrr" => {
            ctx.pending_para_border_target = Some(BorderTarget::Right);
            apply_paragraph_border(ctx);
        }
        "cell" => {
            flush_text(ctx, store, None)?;
            finalize_cell(ctx, store);
            ctx.current_cell = Some(TableCell::new());
        }
        "row" => {
            flush_text(ctx, store, None)?;
            finalize_cell(ctx, store);
            finalize_row(ctx, store);
        }
        _ => return Ok(false),
    }
    Ok(true)
}

fn handle_encoding_controls(word: &str, param: Option<i32>, ctx: &mut RtfParseContext) -> bool {
    match word {
        "ansicpg" => {
            if let Some(cp) = param {
                if let Some(enc) = encoding_for_codepage(cp as u32) {
                    ctx.encoding = enc;
                }
            }
        }
        _ => return false,
    }
    true
}

fn handle_paragraph_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<bool, ParseError> {
    match word {
        "par" => {
            flush_text(ctx, store, None)?;
            if ctx.current_group_kind() == GroupKind::Normal {
                finalize_paragraph(ctx, store);
            }
        }
        "pard" => {
            flush_text(ctx, store, None)?;
            finalize_paragraph(ctx, store);
            ctx.pending_para_style = None;
            ctx.pending_alignment = None;
            ctx.pending_indent = Indentation::default();
            ctx.pending_spacing = Spacing::default();
            ctx.pending_line_rule = None;
            ctx.pending_para_border_target = None;
            ctx.pending_para_borders = ParagraphBorders::default();
        }
        "plain" => {
            ctx.current_props = RtfStyleState::default();
        }
        "line" => {
            append_text(ctx, "\n");
        }
        "tab" => {
            append_text(ctx, "\t");
        }
        "ql" => {
            apply_paragraph_alignment(ctx, TextAlignment::Left);
        }
        "qr" => {
            apply_paragraph_alignment(ctx, TextAlignment::Right);
        }
        "qc" => {
            apply_paragraph_alignment(ctx, TextAlignment::Center);
        }
        "qj" => {
            apply_paragraph_alignment(ctx, TextAlignment::Justify);
        }
        "li" => {
            if let Some(value) = param {
                ctx.pending_indent.left = Some(value);
                if let Some(para) = ctx.current_paragraph.as_mut() {
                    para.properties.indentation = Some(ctx.pending_indent.clone());
                }
            }
        }
        "ri" => {
            if let Some(value) = param {
                ctx.pending_indent.right = Some(value);
                if let Some(para) = ctx.current_paragraph.as_mut() {
                    para.properties.indentation = Some(ctx.pending_indent.clone());
                }
            }
        }
        "fi" => {
            if let Some(value) = param {
                ctx.pending_indent.first_line = Some(value);
                if let Some(para) = ctx.current_paragraph.as_mut() {
                    para.properties.indentation = Some(ctx.pending_indent.clone());
                }
            }
        }
        "sb" => {
            if let Some(value) = param {
                ctx.pending_spacing.before = Some(value.max(0) as u32);
                if let Some(para) = ctx.current_paragraph.as_mut() {
                    para.properties.spacing = Some(ctx.pending_spacing.clone());
                }
            }
        }
        "sa" => {
            if let Some(value) = param {
                ctx.pending_spacing.after = Some(value.max(0) as u32);
                if let Some(para) = ctx.current_paragraph.as_mut() {
                    para.properties.spacing = Some(ctx.pending_spacing.clone());
                }
            }
        }
        "sl" => {
            if let Some(value) = param {
                ctx.pending_spacing.line = Some(value.abs() as u32);
                ctx.pending_spacing.line_rule = ctx.pending_line_rule;
                if let Some(para) = ctx.current_paragraph.as_mut() {
                    para.properties.spacing = Some(ctx.pending_spacing.clone());
                }
            }
        }
        "slmult" => {
            if let Some(value) = param {
                ctx.pending_line_rule = if value == 0 {
                    Some(LineSpacingRule::Exact)
                } else if value == 1 {
                    Some(LineSpacingRule::AtLeast)
                } else {
                    Some(LineSpacingRule::Auto)
                };
            }
        }
        _ => return Ok(false),
    }
    Ok(true)
}

fn apply_paragraph_alignment(ctx: &mut RtfParseContext, alignment: TextAlignment) {
    ctx.pending_alignment = Some(alignment);
    if let Some(para) = ctx.current_paragraph.as_mut() {
        para.properties.alignment = ctx.pending_alignment;
    }
}

fn handle_run_style_controls(word: &str, param: Option<i32>, ctx: &mut RtfParseContext) -> bool {
    match word {
        "b" => {
            ctx.current_props.bold = Some(param.unwrap_or(1) != 0);
        }
        "i" => {
            ctx.current_props.italic = Some(param.unwrap_or(1) != 0);
        }
        "ul" => {
            ctx.current_props.underline = Some(param.unwrap_or(1) != 0);
        }
        "ulnone" => {
            ctx.current_props.underline = Some(false);
        }
        "strike" => {
            ctx.current_props.strike = Some(param.unwrap_or(1) != 0);
        }
        "fs" => {
            if let Some(sz) = param {
                ctx.current_props.font_size = Some(sz.max(0) as u32);
            }
        }
        "f" => {
            if let Some(idx) = param {
                ctx.current_props.font_index = Some(idx.max(0) as u32);
            }
        }
        "cf" => {
            if let Some(idx) = param {
                ctx.current_props.color_index = Some(idx.max(0) as usize);
            }
        }
        "highlight" => {
            if let Some(idx) = param {
                ctx.current_props.highlight_index = Some(idx.max(0) as usize);
            }
        }
        "super" => {
            ctx.current_props.vertical = Some(docir_core::ir::VerticalTextAlignment::Superscript);
        }
        "sub" => {
            ctx.current_props.vertical = Some(docir_core::ir::VerticalTextAlignment::Subscript);
        }
        "nosupersub" => {
            ctx.current_props.vertical = Some(docir_core::ir::VerticalTextAlignment::Baseline);
        }
        _ => return false,
    }
    true
}

fn flush_text(
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
    span: Option<SourceSpan>,
) -> Result<(), ParseError> {
    if ctx.current_text.is_empty() {
        return Ok(());
    }

    let text = ctx.current_text.clone();
    ctx.current_text.clear();

    if ctx.current_group_kind() == GroupKind::FontTable {
        parse_font_entry(&text, ctx);
        return Ok(());
    }
    if ctx.current_group_kind() == GroupKind::ColorTable {
        parse_color_entries(&text, ctx);
        return Ok(());
    }
    if matches!(
        ctx.current_group_kind(),
        GroupKind::Stylesheet | GroupKind::StylesheetEntry
    ) {
        if let Some(mut style_ctx) = ctx.current_style.take() {
            let mut pending = format!("{}{}", style_ctx.name_buf, text);
            loop {
                if let Some(pos) = pending.find(';') {
                    let (head, rest) = pending.split_at(pos);
                    let name = head.trim().to_string();
                    push_style_from_ctx(ctx, &style_ctx, name);
                    pending = rest.trim_start_matches(';').to_string();
                    style_ctx.name_buf.clear();
                } else {
                    style_ctx.name_buf.push_str(pending.trim());
                    ctx.current_style = Some(style_ctx);
                    break;
                }
            }
        }
        return Ok(());
    }
    if ctx.current_group_kind() == GroupKind::Object {
        if let Some(target) = ctx.object_text_target {
            if let Some(obj) = ctx.object_stack.last_mut() {
                match target {
                    ObjectTextTarget::Class => obj.class_name = Some(text.trim().to_string()),
                    ObjectTextTarget::Name => obj.object_name = Some(text.trim().to_string()),
                }
            }
        }
        return Ok(());
    }

    let props = run_properties_from_state(ctx);
    let mut run = Run::with_properties(text.clone(), props);
    run.span = span.clone();
    let run_id = run.id;
    store.insert(IRNode::Run(run));

    match ctx.current_group_kind() {
        GroupKind::FieldInst => {
            if let Some(field) = ctx.field_stack.last_mut() {
                field.instruction.push_str(&text);
            }
        }
        GroupKind::FieldResult => {
            if let Some(field) = ctx.field_stack.last_mut() {
                field.runs.push(run_id);
            }
        }
        _ => {
            ensure_paragraph(ctx, store);
            if let Some(para) = ctx.current_paragraph.as_mut() {
                para.runs.push(run_id);
            }
        }
    }

    Ok(())
}

fn append_text(ctx: &mut RtfParseContext, text: &str) {
    ctx.current_text.push_str(text);
}

fn append_text_byte(ctx: &mut RtfParseContext, byte: u8) {
    let binding = [byte];
    let (text, _, _) = ctx.encoding.decode(&binding);
    ctx.current_text.push_str(&text);
}

fn ensure_paragraph(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if ctx.current_paragraph.is_none() {
        let mut para = Paragraph::new();
        para.span = Some(SourceSpan::new("rtf"));
        apply_pending_paragraph(&mut para, ctx);
        ctx.current_paragraph = Some(para);
    }
    ensure_section(ctx, store);
}

fn ensure_section(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if ctx.current_section.is_none() {
        let mut section = docir_core::ir::Section::new();
        section.span = Some(SourceSpan::new("rtf"));
        let section_id = section.id;
        store.insert(IRNode::Section(section));
        ctx.current_section = Some(section_id);
        ctx.sections.push(section_id);
    }
}

fn apply_pending_paragraph(para: &mut Paragraph, ctx: &mut RtfParseContext) {
    if let Some(style_id) = ctx.pending_para_style.clone() {
        para.style_id = Some(style_id);
    }
    if let Some(numbering) = pending_numbering(ctx) {
        para.properties.numbering = Some(numbering);
    }
    if let Some(align) = ctx.pending_alignment {
        para.properties.alignment = Some(align);
    }
    if ctx.pending_indent.left.is_some()
        || ctx.pending_indent.right.is_some()
        || ctx.pending_indent.first_line.is_some()
        || ctx.pending_indent.hanging.is_some()
    {
        para.properties.indentation = Some(ctx.pending_indent.clone());
    }
    if ctx.pending_spacing.before.is_some()
        || ctx.pending_spacing.after.is_some()
        || ctx.pending_spacing.line.is_some()
        || ctx.pending_spacing.line_rule.is_some()
    {
        para.properties.spacing = Some(ctx.pending_spacing.clone());
    }
    if let Some(borders) = pending_paragraph_borders(ctx) {
        para.properties.borders = Some(borders);
    }
}

fn pending_numbering(ctx: &RtfParseContext) -> Option<NumberingInfo> {
    let list_override = ctx.pending_list_override?;
    let level = ctx.pending_list_level.unwrap_or(0);
    let list_id = ctx
        .list_overrides
        .get(&list_override)
        .copied()
        .unwrap_or(list_override);
    let format = ctx.list_level_formats.get(&(list_id, level)).cloned();
    Some(NumberingInfo {
        num_id: list_id as u32,
        level,
        format,
    })
}

fn color_from_index(ctx: &RtfParseContext, index: usize) -> Option<String> {
    ctx.color_table.colors.get(index).and_then(|c| c.clone())
}

fn set_border_target(ctx: &mut RtfParseContext, target: BorderTarget) {
    ctx.pending_border_target = Some(target);
}

fn apply_border(ctx: &mut RtfParseContext) {
    if ctx.pending_para_border_target.is_some() {
        ctx.pending_para_border = ctx.pending_border.clone();
        apply_paragraph_border(ctx);
        return;
    }
    let Some(target) = ctx.pending_border_target else {
        return;
    };
    let props = ctx
        .pending_cell_props
        .get_or_insert_with(TableCellProperties::default);
    let mut borders = props.borders.take().unwrap_or_default();
    let border = ctx.pending_border.clone();
    match target {
        BorderTarget::Top => borders.top = Some(border),
        BorderTarget::Bottom => borders.bottom = Some(border),
        BorderTarget::Left => borders.left = Some(border),
        BorderTarget::Right => borders.right = Some(border),
        BorderTarget::InsideH => borders.inside_h = Some(border),
        BorderTarget::InsideV => borders.inside_v = Some(border),
    }
    props.borders = Some(borders);
}

fn apply_paragraph_border(ctx: &mut RtfParseContext) {
    let Some(target) = ctx.pending_para_border_target else {
        return;
    };
    let border = ctx.pending_para_border.clone();
    if let Some(para) = ctx.current_paragraph.as_mut() {
        let mut borders = para.properties.borders.take().unwrap_or_default();
        match target {
            BorderTarget::Top => borders.top = Some(border),
            BorderTarget::Bottom => borders.bottom = Some(border),
            BorderTarget::Left => borders.left = Some(border),
            BorderTarget::Right => borders.right = Some(border),
            BorderTarget::InsideH | BorderTarget::InsideV => {}
        }
        para.properties.borders = Some(borders);
    } else {
        match target {
            BorderTarget::Top => ctx.pending_para_borders.top = Some(border),
            BorderTarget::Bottom => ctx.pending_para_borders.bottom = Some(border),
            BorderTarget::Left => ctx.pending_para_borders.left = Some(border),
            BorderTarget::Right => ctx.pending_para_borders.right = Some(border),
            BorderTarget::InsideH | BorderTarget::InsideV => {}
        }
    }
}

fn pending_paragraph_borders(ctx: &RtfParseContext) -> Option<ParagraphBorders> {
    let borders = &ctx.pending_para_borders;
    if borders.top.is_none()
        && borders.bottom.is_none()
        && borders.left.is_none()
        && borders.right.is_none()
    {
        None
    } else {
        Some(borders.clone())
    }
}

fn cell_width_from_row(ctx: &RtfParseContext) -> Option<TableWidth> {
    if ctx.row_cellx.is_empty() {
        return None;
    }
    let idx = ctx.current_cell_index;
    let end = *ctx.row_cellx.get(idx)?;
    let start = if idx == 0 {
        0
    } else {
        *ctx.row_cellx.get(idx - 1)?
    };
    let width = (end - start).max(0) as u32;
    Some(TableWidth {
        value: width,
        width_type: TableWidthType::Dxa,
    })
}

fn finalize_paragraph(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if let Some(para) = ctx.current_paragraph.take() {
        let para_id = para.id;
        store.insert(IRNode::Paragraph(para));
        if let Some(cell) = ctx.current_cell.as_mut() {
            cell.content.push(para_id);
        } else if let Some(section_id) = ctx.current_section {
            if let Some(IRNode::Section(section)) = store.get_mut(section_id) {
                section.content.push(para_id);
            }
        }
    }
    ctx.pending_list_override = None;
    ctx.pending_list_level = None;
}

fn finalize_section(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if let Some(section_id) = ctx.current_section.take() {
        if store.get(section_id).is_none() {
            let mut section = docir_core::ir::Section::new();
            section.id = section_id;
            section.span = Some(SourceSpan::new("rtf"));
            store.insert(IRNode::Section(section));
        }
    }
}

fn finalize_cell(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if let Some(mut cell) = ctx.current_cell.take() {
        if let Some(mut props) = ctx.pending_cell_props.take() {
            if let Some(width) = cell_width_from_row(ctx) {
                props.width = Some(width);
            }
            cell.properties = props;
        } else if let Some(width) = cell_width_from_row(ctx) {
            cell.properties.width = Some(width);
        }
        ctx.current_cell_index = ctx.current_cell_index.saturating_add(1);
        let cell_id = cell.id;
        store.insert(IRNode::TableCell(cell));
        if let Some(row) = ctx.current_row.as_mut() {
            row.cells.push(cell_id);
        }
    }
}

fn finalize_row(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if let Some(row) = ctx.current_row.take() {
        let row_id = row.id;
        store.insert(IRNode::TableRow(row));
        if let Some(table) = ctx.current_table.as_mut() {
            table.rows.push(row_id);
        }
    }
    ctx.row_cellx.clear();
    ctx.current_cell_index = 0;
}

fn finalize_table_if_open(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if let Some(table) = ctx.current_table.take() {
        let table_id = table.id;
        store.insert(IRNode::Table(table));
        ensure_section(ctx, store);
        if let Some(section_id) = ctx.current_section {
            if let Some(IRNode::Section(section)) = store.get_mut(section_id) {
                section.content.push(table_id);
            }
        }
    }
}

fn finalize_field(ctx: &mut RtfParseContext, store: &mut IrStore) {
    let Some(field) = ctx.field_stack.pop() else {
        return;
    };
    let instr = field.instruction.trim();
    let mut instruction = if instr.is_empty() {
        None
    } else {
        Some(instr.to_string())
    };
    if let Some(instr_text) = instruction.clone() {
        if let Some((target, _args, _switches)) = parse_hyperlink_instruction(&instr_text) {
            let mut link = Hyperlink::new(target, true);
            link.runs = field.runs.clone();
            let link_id = link.id;
            store.insert(IRNode::Hyperlink(link));
            ensure_paragraph(ctx, store);
            if let Some(para) = ctx.current_paragraph.as_mut() {
                para.runs.push(link_id);
            }
            if let Some(ext_id) = create_external_ref(&instr_text, store, ctx) {
                ctx.external_refs.push(ext_id);
            }
            return;
        }
    }

    let mut node = Field::new(instruction.take());
    node.runs = field.runs.clone();
    if let Some(instr_text) = node.instruction.as_ref() {
        node.instruction_parsed = parse_field_instruction(instr_text);
    }
    let field_id = node.id;
    store.insert(IRNode::Field(node));
    ensure_paragraph(ctx, store);
    if let Some(para) = ctx.current_paragraph.as_mut() {
        para.runs.push(field_id);
    }
}

fn finalize_style_entry(ctx: &mut RtfParseContext) {
    let Some(style_ctx) = ctx.current_style.take() else {
        return;
    };
    let name = style_ctx.name_buf.trim().to_string();
    if !name.is_empty() {
        push_style_from_ctx(ctx, &style_ctx, name);
    }
}

fn push_style_from_ctx(ctx: &mut RtfParseContext, style_ctx: &StyleEntryContext, name: String) {
    let style_id = style_ctx
        .style_id
        .clone()
        .unwrap_or_else(|| "style".to_string());
    let style_type = style_ctx.style_type.unwrap_or(StyleType::Other);
    let style = Style {
        style_id,
        name: if name.is_empty() { None } else { Some(name) },
        style_type,
        based_on: None,
        next: None,
        is_default: false,
        run_props: None,
        paragraph_props: None,
        table_props: None,
    };
    if ctx.style_set.is_none() {
        ctx.style_set = Some(StyleSet::new());
    }
    if let Some(set) = ctx.style_set.as_mut() {
        set.styles.push(style);
    }
}

fn finalize_list_entry(ctx: &mut RtfParseContext) {
    if let Some(list_id) = ctx.current_list_id {
        let levels = ctx.current_list_level.max(1);
        ctx.list_levels.insert(list_id, levels);
    }
    ctx.current_list_id = None;
    ctx.current_list_level = 0;
}

fn finalize_list_override(ctx: &mut RtfParseContext) {
    if let (Some(override_id), Some(list_id)) =
        (ctx.current_list_override, ctx.current_list_override_list_id)
    {
        ctx.list_overrides.insert(override_id, list_id);
    }
    ctx.current_list_override = None;
    ctx.current_list_override_list_id = None;
}

fn parse_field_instruction(text: &str) -> Option<FieldInstruction> {
    let tokens = tokenize_field_instruction(text);
    if tokens.is_empty() {
        return None;
    }
    let first = tokens[0].to_ascii_uppercase();
    let kind = match first.as_str() {
        "HYPERLINK" => FieldKind::Hyperlink,
        "INCLUDETEXT" => FieldKind::IncludeText,
        "MERGEFIELD" => FieldKind::MergeField,
        "DATE" => FieldKind::Date,
        "REF" => FieldKind::Ref,
        "PAGEREF" => FieldKind::PageRef,
        _ => FieldKind::Unknown,
    };
    let mut args = Vec::new();
    let mut switches = Vec::new();
    for token in tokens.iter().skip(1) {
        if token.starts_with('\\') {
            switches.push(token.trim_start_matches('\\').to_string());
        } else {
            args.push(token.to_string());
        }
    }
    Some(FieldInstruction {
        kind,
        args,
        switches,
    })
}

fn parse_hyperlink_instruction(text: &str) -> Option<(String, Vec<String>, Vec<String>)> {
    let tokens = tokenize_field_instruction(text);
    if tokens.is_empty() || tokens[0].to_ascii_uppercase() != "HYPERLINK" {
        return None;
    }
    let mut target = None;
    let mut args = Vec::new();
    let mut switches = Vec::new();
    for token in tokens.into_iter().skip(1) {
        if token.starts_with('\\') {
            switches.push(token.trim_start_matches('\\').to_string());
        } else if target.is_none() {
            target = Some(token);
        } else {
            args.push(token);
        }
    }
    target.map(|t| (t, args, switches))
}

fn tokenize_field_instruction(text: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    let mut current = String::new();
    let mut in_quotes = false;
    let mut chars = text.chars().peekable();
    while let Some(ch) = chars.next() {
        match ch {
            '"' => {
                in_quotes = !in_quotes;
            }
            '\\' => {
                if !current.is_empty() {
                    tokens.push(current.clone());
                    current.clear();
                }
                let mut switch = String::from("\\");
                while let Some(&c) = chars.peek() {
                    if c.is_whitespace() {
                        break;
                    }
                    if c == '"' {
                        break;
                    }
                    switch.push(c);
                    chars.next();
                }
                tokens.push(switch);
            }
            c if c.is_whitespace() && !in_quotes => {
                if !current.is_empty() {
                    tokens.push(current.clone());
                    current.clear();
                }
            }
            _ => current.push(ch),
        }
    }
    if !current.is_empty() {
        tokens.push(current);
    }
    tokens
}

fn create_external_ref(
    instr: &str,
    store: &mut IrStore,
    _ctx: &mut RtfParseContext,
) -> Option<NodeId> {
    if let Some((target, _, _)) = parse_hyperlink_instruction(instr) {
        let mut ext = ExternalReference::new(ExternalRefType::Hyperlink, target.clone());
        ext.span = Some(SourceSpan::new("rtf"));
        let id = ext.id;
        store.insert(IRNode::ExternalReference(ext));
        return Some(id);
    }
    None
}

fn run_properties_from_state(ctx: &RtfParseContext) -> RunProperties {
    let mut props = RunProperties::default();
    props.style_id = ctx.current_props.style_id.clone();
    props.bold = ctx.current_props.bold;
    props.italic = ctx.current_props.italic;
    props.underline = ctx.current_props.underline.map(|u| {
        if u {
            docir_core::ir::UnderlineStyle::Single
        } else {
            docir_core::ir::UnderlineStyle::None
        }
    });
    props.strike = ctx.current_props.strike;
    props.font_size = ctx.current_props.font_size;
    props.vertical_align = ctx.current_props.vertical;
    if let Some(idx) = ctx.current_props.font_index {
        if let Some(name) = ctx.font_table.fonts.get(&idx) {
            props.font_family = Some(name.clone());
        }
    }
    if let Some(idx) = ctx.current_props.color_index {
        if let Some(color) = ctx.color_table.colors.get(idx).and_then(|c| c.clone()) {
            props.color = Some(color);
        }
    }
    if let Some(idx) = ctx.current_props.highlight_index {
        if let Some(color) = ctx.color_table.colors.get(idx).and_then(|c| c.clone()) {
            props.highlight = Some(color);
        }
    }
    props
}

fn parse_font_entry(text: &str, ctx: &mut RtfParseContext) {
    // Parse entries like "\\f0\\fnil Helvetica;"
    let mut font_index: Option<u32> = None;
    let mut name = String::new();
    let mut chars = text.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '\\' {
            let mut control = String::new();
            while let Some(c) = chars.peek() {
                if c.is_alphanumeric() {
                    control.push(*c);
                    chars.next();
                } else {
                    break;
                }
            }
            if control.starts_with('f') {
                let num = control.trim_start_matches('f');
                if let Ok(idx) = num.parse::<u32>() {
                    font_index = Some(idx);
                }
            }
        } else if ch == ';' {
            break;
        } else if !ch.is_control() {
            name.push(ch);
        }
    }
    let name = name.trim().to_string();
    if let Some(idx) = font_index {
        if !name.is_empty() {
            ctx.font_table.fonts.insert(idx, name);
        }
    }
}

fn parse_color_entries(text: &str, ctx: &mut RtfParseContext) {
    let mut red: Option<u8> = None;
    let mut green: Option<u8> = None;
    let mut blue: Option<u8> = None;
    let mut chars = text.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '\\' {
            let mut control = String::new();
            while let Some(c) = chars.peek() {
                if c.is_alphanumeric() {
                    control.push(*c);
                    chars.next();
                } else {
                    break;
                }
            }
            if control.starts_with('r') {
                red = control[1..].parse::<u8>().ok();
            } else if control.starts_with('g') {
                green = control[1..].parse::<u8>().ok();
            } else if control.starts_with('b') {
                blue = control[1..].parse::<u8>().ok();
            }
        } else if ch == ';' {
            let color = match (red, green, blue) {
                (Some(r), Some(g), Some(b)) => Some(format!("{:02X}{:02X}{:02X}", r, g, b)),
                _ => None,
            };
            ctx.color_table.colors.push(color);
            red = None;
            green = None;
            blue = None;
        }
    }
}

fn hex_val(b: u8) -> Option<u8> {
    match b {
        b'0'..=b'9' => Some(b - b'0'),
        b'a'..=b'f' => Some(10 + b - b'a'),
        b'A'..=b'F' => Some(10 + b - b'A'),
        _ => None,
    }
}

fn encoding_for_codepage(cp: u32) -> Option<&'static Encoding> {
    match cp {
        65001 => Some(encoding_rs::UTF_8),
        1252 => Some(encoding_rs::WINDOWS_1252),
        1250 => Some(encoding_rs::WINDOWS_1250),
        1251 => Some(encoding_rs::WINDOWS_1251),
        1253 => Some(encoding_rs::WINDOWS_1253),
        1254 => Some(encoding_rs::WINDOWS_1254),
        1255 => Some(encoding_rs::WINDOWS_1255),
        1256 => Some(encoding_rs::WINDOWS_1256),
        1257 => Some(encoding_rs::WINDOWS_1257),
        1258 => Some(encoding_rs::WINDOWS_1258),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::error::ParseError;
    use crate::parser::ParserConfig;
    use docir_core::types::NodeType;

    #[test]
    fn parse_simple_rtf() {
        let data = b"{\\rtf1\\ansi Hello \\par World}";
        let parser = RtfParser::new();
        let parsed = parser.parse_bytes(data).expect("parse rtf");
        assert_eq!(parsed.format, DocumentFormat::Rtf);
    }

    #[test]
    fn parse_hyperlink_field() {
        let data =
            b"{\\rtf1{\\field{\\fldinst HYPERLINK \\\"https://example.com\\\"}{\\fldrslt Link}}}";
        let parser = RtfParser::new();
        let parsed = parser.parse_bytes(data).expect("parse rtf");
        assert_eq!(parsed.format, DocumentFormat::Rtf);
    }

    #[test]
    fn parse_styles_and_lists() {
        let data = b"{\\rtf1\\ansi{\\stylesheet{\\s1 Heading 1;}{\\cs2 Emphasis;}}\\pard\\ql\\s1\\ls1\\ilvl0 Item}";
        let parser = RtfParser::new();
        let parsed = parser.parse_bytes(data).expect("parse rtf");
        let doc = parsed.document().expect("doc");
        assert!(doc.styles.is_some());
        let has_style_set = parsed
            .store
            .iter_ids_by_type(NodeType::StyleSet)
            .next()
            .is_some();
        assert!(has_style_set);
        let has_numbering = parsed.store.values().any(|node| match node {
            IRNode::Paragraph(p) => p.properties.numbering.is_some(),
            _ => false,
        });
        assert!(has_numbering);
    }

    #[test]
    fn parse_table_borders_and_widths() {
        let data = b"{\\rtf1\\ansi{\\colortbl;\\red255\\green0\\blue0;}\\trowd\\cellx1000\\cellx2000\\clbrdrt\\brdrs\\brdrw10\\clcbpat1\\cell One\\cell Two\\row}";
        let parser = RtfParser::new();
        let parsed = parser.parse_bytes(data).expect("parse rtf");
        let has_cell_props = parsed.store.values().any(|node| match node {
            IRNode::TableCell(cell) => {
                cell.properties.width.is_some()
                    || cell.properties.borders.is_some()
                    || cell.properties.shading.is_some()
            }
            _ => false,
        });
        assert!(has_cell_props);
    }

    #[test]
    fn parse_paragraph_margins_and_borders() {
        let data = b"{\\rtf1\\ansi\\pard\\li720\\ri360\\fi180\\sb120\\sa240\\sl360\\slmult1\\brdrt\\brdrs\\brdrw15 Paragraph}";
        let parser = RtfParser::new();
        let parsed = parser.parse_bytes(data).expect("parse rtf");
        let has_props = parsed.store.values().any(|node| match node {
            IRNode::Paragraph(p) => {
                p.properties.indentation.is_some()
                    || p.properties.spacing.is_some()
                    || p.properties.borders.is_some()
            }
            _ => false,
        });
        assert!(has_props);
    }

    #[test]
    fn rtf_group_depth_limit() {
        let data = b"{\\rtf1{{{a}}}}";
        let mut config = ParserConfig::default();
        config.rtf_max_group_depth = 3;
        let parser = RtfParser::with_config(config);
        let err = parser
            .parse_bytes(data)
            .expect_err("should hit depth limit");
        match err {
            ParseError::ResourceLimit(message) => {
                assert!(message.contains("RTF max group depth"));
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn rtf_object_hex_limit() {
        let data = b"{\\rtf1{\\object{\\objdata 0102030405}}}";
        let mut config = ParserConfig::default();
        config.rtf_max_object_hex_len = 4;
        let parser = RtfParser::with_config(config);
        let err = parser
            .parse_bytes(data)
            .expect_err("should hit objdata limit");
        match err {
            ParseError::ResourceLimit(message) => {
                assert!(message.contains("RTF objdata too large"));
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }
}
