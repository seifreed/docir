//! ODF (OpenDocument) parsing support.

use crate::diagnostics::{attach_diagnostics_if_any, push_entry, push_info, push_warning};
use crate::error::ParseError;
use crate::format::FormatParser;
use crate::input::enforce_input_size;
use crate::parser::{ParsedDocument, ParserConfig};
use crate::text_utils::parse_text_alignment;
use crate::xml_utils::{attr_value, read_event};
use crate::zip_handler::SecureZipReader;
use docir_core::ir::{
    BookmarkEnd, BookmarkStart, Cell, CellFormula, CellValue, ChartData, Comment, CommentReference,
    ConditionalFormat, ConditionalRule, DataValidation, DefinedName, DiagnosticSeverity,
    Diagnostics, Document, Endnote, ExtensionPart, ExtensionPartKind, Field, FieldInstruction,
    FieldKind, Footer, Footnote, Header, IRNode, MediaAsset, MediaType, MergedCellRange,
    NumberingInfo, Paragraph, ParagraphProperties, PivotCache, PivotCacheRecords, PivotTable,
    Revision, RevisionType, Run, Section, Shape, ShapeText, ShapeTextParagraph, ShapeTextRun,
    ShapeTransform, ShapeType, Slide, SlideAnimation, SlideTransition, Style, StyleSet, StyleType,
    Table, TableCell, TableCellProperties, TableRow, Worksheet, WorksheetDrawing,
};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::HashMap;
use std::io::{Read, Seek};
use std::sync::Arc;
use std::thread;

mod builder;
mod container;
mod formula;
mod io;
mod limits;
mod manifest;
mod ods;
mod paragraph;
mod presentation;
mod presentation_helpers;
pub(crate) mod security;
mod security_helpers;
mod spreadsheet;
mod text;

use self::security::scan_odf_filters;
use self::security_helpers::{build_odf_macro_project, parse_odf_signatures};
use crate::security_scan::{DefaultSecurityScanner, SecurityScanner};

use container::{handle_content_xml, load_meta};
use io::{collect_manifest_index, collect_shared_parts};
use limits::{OdfAtomicLimits, OdfLimitCounter, OdfLimits};
use manifest::{
    encrypted_manifest_entries, format_odf_encryption_metadata, is_manifest_entry_encrypted,
    parse_manifest, OdfEncryptionData, OdfManifestEntry,
};
use ods::{parse_ods_cell, parse_ods_cell_empty, parse_ods_table, parse_ods_table_fast};
use paragraph::parse_paragraph;
use presentation_helpers::{
    build_media_asset, classify_media_shape, parse_draw_page, parse_odf_chart, parse_odp_transition,
};

type OdfReader<'a> = Reader<std::io::Cursor<&'a [u8]>>;

/// ODF parser (ODT/ODS/ODP).
pub struct OdfParser {
    config: ParserConfig,
}

impl FormatParser for OdfParser {
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        self.parse_reader(reader)
    }
}

#[derive(Debug, Default)]
struct OdfContentResult {
    content: Vec<NodeId>,
    comments: Vec<NodeId>,
    footnotes: Vec<NodeId>,
    endnotes: Vec<NodeId>,
    pivot_caches: Vec<NodeId>,
}

impl OdfParser {
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

    crate::impl_parse_entrypoints!();
}

fn parse_content(
    xml: &[u8],
    format: DocumentFormat,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<OdfContentResult, ParseError> {
    match format {
        DocumentFormat::OdfText => text::parse_content_text(xml, store, limits),
        DocumentFormat::OdfSpreadsheet => {
            spreadsheet::parse_content_spreadsheet(xml, store, limits)
        }
        DocumentFormat::OdfPresentation => {
            presentation::parse_content_presentation(xml, store, limits)
        }
        _ => Ok(OdfContentResult::default()),
    }
}

fn strip_odf_formula_prefix(formula: &str) -> &str {
    if let Some(stripped) = formula.strip_prefix("of:=") {
        stripped
    } else if let Some(stripped) = formula.strip_prefix("of:") {
        stripped
    } else {
        formula
    }
}

fn parse_ods_named_ranges(xml: &[u8]) -> Vec<DefinedName> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:named-range" => {
                    if let Some(name) = attr_value(&e, b"table:name") {
                        let value = attr_value(&e, b"table:cell-range-address")
                            .unwrap_or_else(|| String::new());
                        let mut def = DefinedName {
                            id: NodeId::new(),
                            name,
                            value,
                            local_sheet_id: None,
                            hidden: false,
                            comment: attr_value(&e, b"table:comment"),
                            span: Some(SourceSpan::new("content.xml")),
                        };
                        if let Some(hidden) = attr_value(&e, b"table:hidden") {
                            def.hidden = hidden == "true";
                        }
                        out.push(def);
                    }
                }
                b"table:named-expression" => {
                    if let Some(name) = attr_value(&e, b"table:name") {
                        let value =
                            attr_value(&e, b"table:expression").unwrap_or_else(|| String::new());
                        let mut def = DefinedName {
                            id: NodeId::new(),
                            name,
                            value,
                            local_sheet_id: None,
                            hidden: false,
                            comment: attr_value(&e, b"table:comment"),
                            span: Some(SourceSpan::new("content.xml")),
                        };
                        if let Some(hidden) = attr_value(&e, b"table:hidden") {
                            def.hidden = hidden == "true";
                        }
                        out.push(def);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    out
}

fn parse_frame_transform(start: &BytesStart<'_>) -> ShapeTransform {
    let mut transform = ShapeTransform::default();
    if let Some(x) = attr_value(start, b"svg:x").and_then(parse_length_emu) {
        transform.x = x;
    }
    if let Some(y) = attr_value(start, b"svg:y").and_then(parse_length_emu) {
        transform.y = y;
    }
    if let Some(width) = attr_value(start, b"svg:width").and_then(parse_length_emu_u64) {
        transform.width = width;
    }
    if let Some(height) = attr_value(start, b"svg:height").and_then(parse_length_emu_u64) {
        transform.height = height;
    }
    transform
}

fn parse_length_emu(value: String) -> Option<i64> {
    parse_length_emu_str(&value).map(|v| v.round() as i64)
}

fn parse_length_emu_u64(value: String) -> Option<u64> {
    parse_length_emu_str(&value).map(|v| v.max(0.0).round() as u64)
}

fn parse_length_emu_str(value: &str) -> Option<f64> {
    let trimmed = value.trim();
    let mut num = String::new();
    let mut unit = String::new();
    for ch in trimmed.chars() {
        if ch.is_ascii_digit() || ch == '.' || ch == '-' {
            num.push(ch);
        } else {
            unit.push(ch);
        }
    }
    let magnitude = num.parse::<f64>().ok()?;
    let emu = match unit.as_str() {
        "cm" => magnitude / 2.54 * 914_400.0,
        "mm" => magnitude / 25.4 * 914_400.0,
        "in" => magnitude * 914_400.0,
        "pt" => magnitude * 12_700.0,
        "pc" => magnitude * 152_400.0,
        "px" => magnitude * 9_525.0,
        "" => magnitude,
        _ => return None,
    };
    Some(emu)
}

fn parse_draw_frame_presentation(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let transform = parse_frame_transform(start);
    let mut shape_type = ShapeType::Picture;
    let mut media_target: Option<String> = None;
    let mut text: Option<ShapeText> = None;
    let mut name = attr_value(start, b"draw:name");
    let mut chart_id: Option<NodeId> = None;
    let mut buf = Vec::new();
    let mut has_shape = false;

    loop {
        match read_event(reader, &mut buf, "content.xml")? {
            Event::Start(e) => match e.name().as_ref() {
                b"draw:text-box" => {
                    let paragraphs = parse_shape_text(reader, b"draw:text-box")?;
                    if !paragraphs.is_empty() {
                        text = Some(ShapeText { paragraphs });
                        shape_type = ShapeType::TextBox;
                        has_shape = true;
                    }
                }
                b"draw:image" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href);
                        shape_type = ShapeType::Picture;
                        has_shape = true;
                    }
                }
                b"chart:chart" => {
                    shape_type = ShapeType::Chart;
                    has_shape = true;
                    let chart = parse_odf_chart(reader, &e)?;
                    let id = chart.id;
                    store.insert(IRNode::ChartData(chart));
                    chart_id = Some(id);
                }
                b"draw:plugin" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                        shape_type = classify_media_shape(&href);
                        has_shape = true;
                    }
                }
                b"draw:object" | b"draw:object-ole" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                    }
                    shape_type = ShapeType::OleObject;
                    has_shape = true;
                }
                _ => {}
            },
            Event::Empty(e) => match e.name().as_ref() {
                b"draw:image" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href);
                        shape_type = ShapeType::Picture;
                        has_shape = true;
                    }
                }
                b"chart:chart" => {
                    shape_type = ShapeType::Chart;
                    has_shape = true;
                    let mut chart = ChartData::new();
                    chart.chart_type = attr_value(&e, b"chart:class");
                    chart.span = Some(SourceSpan::new("content.xml"));
                    let id = chart.id;
                    store.insert(IRNode::ChartData(chart));
                    chart_id = Some(id);
                }
                b"draw:plugin" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                        shape_type = classify_media_shape(&href);
                        has_shape = true;
                    }
                }
                b"draw:object" | b"draw:object-ole" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                    }
                    shape_type = ShapeType::OleObject;
                    has_shape = true;
                }
                _ => {}
            },
            Event::End(e) => {
                if e.name().as_ref() == b"draw:frame" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if has_shape {
        let mut shape = Shape::new(shape_type);
        shape.name = name.take();
        shape.media_target = media_target;
        shape.text = text;
        shape.chart_id = chart_id;
        shape.transform = transform;
        let shape_id = shape.id;
        store.insert(IRNode::Shape(shape));
        Ok(Some(shape_id))
    } else {
        Ok(None)
    }
}

fn parse_custom_shape_presentation(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let mut name = attr_value(start, b"draw:name");
    let paragraphs = parse_shape_text(reader, b"draw:custom-shape")?;
    let mut shape = Shape::new(ShapeType::Custom);
    shape.name = name.take();
    if !paragraphs.is_empty() {
        shape.text = Some(ShapeText { paragraphs });
    }
    let shape_id = shape.id;
    store.insert(IRNode::Shape(shape));
    Ok(Some(shape_id))
}

fn parse_shape_text(
    reader: &mut OdfReader<'_>,
    end_tag: &[u8],
) -> Result<Vec<ShapeTextParagraph>, ParseError> {
    let mut buf = Vec::new();
    let mut paragraphs = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:p" => {
                    let text = parse_text_element(reader, b"text:p")?;
                    let run = ShapeTextRun {
                        text,
                        bold: None,
                        italic: None,
                        font_size: None,
                        font_family: None,
                    };
                    paragraphs.push(ShapeTextParagraph {
                        runs: vec![run],
                        alignment: None,
                    });
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_tag {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(paragraphs)
}

fn parse_notes(reader: &mut OdfReader<'_>) -> Result<Option<String>, ParseError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match read_event(reader, &mut buf, "content.xml")? {
            Event::Start(e) => {
                if e.name().as_ref() == b"text:p" {
                    let para = parse_text_element(reader, b"text:p")?;
                    if !text.is_empty() {
                        text.push('\n');
                    }
                    text.push_str(&para);
                }
            }
            Event::End(e) => {
                if e.name().as_ref() == b"presentation:notes" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if text.is_empty() {
        Ok(None)
    } else {
        Ok(Some(text))
    }
}
#[derive(Debug, Clone)]
struct ValidationDef {
    validation_type: Option<String>,
    operator: Option<String>,
    allow_blank: bool,
    show_input_message: bool,
    show_error_message: bool,
    error_title: Option<String>,
    error: Option<String>,
    prompt_title: Option<String>,
    prompt: Option<String>,
    formula1: Option<String>,
    formula2: Option<String>,
}

fn parse_validation_definition(start: &BytesStart<'_>) -> Option<(String, ValidationDef)> {
    let name = attr_value(start, b"table:name")?;
    let condition = attr_value(start, b"table:condition");
    let allow_blank = attr_value(start, b"table:allow-empty-cell")
        .map(|v| v == "true")
        .unwrap_or(false);
    let show_input_message = attr_value(start, b"table:display-list")
        .map(|v| v == "true")
        .unwrap_or(false);
    let show_error_message = attr_value(start, b"table:display-list")
        .map(|v| v == "true")
        .unwrap_or(false);
    let def = ValidationDef {
        validation_type: condition.clone(),
        operator: None,
        allow_blank,
        show_input_message,
        show_error_message,
        error_title: None,
        error: None,
        prompt_title: None,
        prompt: None,
        formula1: condition,
        formula2: None,
    };
    Some((name, def))
}

fn parse_ods_conditional_formatting(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
) -> Result<Option<ConditionalFormat>, ParseError> {
    let mut cf = ConditionalFormat {
        id: NodeId::new(),
        ranges: Vec::new(),
        rules: Vec::new(),
        span: Some(SourceSpan::new("content.xml")),
    };
    if let Some(ranges) = attr_value(start, b"table:target-range-address")
        .or_else(|| attr_value(start, b"table:cell-range-address"))
    {
        cf.ranges = ranges.split_whitespace().map(|s| s.to_string()).collect();
    }

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"table:conditional-format" {
                    let rule = build_ods_conditional_rule(&e);
                    cf.rules.push(rule);
                    spreadsheet::skip_element(reader, e.name().as_ref())?;
                }
            }
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"table:conditional-format" {
                    let rule = build_ods_conditional_rule(&e);
                    cf.rules.push(rule);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:conditional-formatting" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if cf.rules.is_empty() && cf.ranges.is_empty() {
        Ok(None)
    } else {
        Ok(Some(cf))
    }
}

fn parse_ods_conditional_formatting_empty(
    start: &BytesStart<'_>,
) -> Result<Option<ConditionalFormat>, ParseError> {
    let mut cf = ConditionalFormat {
        id: NodeId::new(),
        ranges: Vec::new(),
        rules: Vec::new(),
        span: Some(SourceSpan::new("content.xml")),
    };
    if let Some(ranges) = attr_value(start, b"table:target-range-address")
        .or_else(|| attr_value(start, b"table:cell-range-address"))
    {
        cf.ranges = ranges.split_whitespace().map(|s| s.to_string()).collect();
    }
    if cf.rules.is_empty() && cf.ranges.is_empty() {
        Ok(None)
    } else {
        Ok(Some(cf))
    }
}

fn build_ods_conditional_rule(start: &BytesStart<'_>) -> ConditionalRule {
    let mut rule = ConditionalRule {
        rule_type: "odf-condition".to_string(),
        priority: None,
        operator: None,
        formulae: Vec::new(),
    };
    rule.priority = attr_value(start, b"table:priority").and_then(|v| v.parse::<u32>().ok());
    if let Some(condition) = attr_value(start, b"table:condition") {
        rule.operator = parse_odf_condition_operator(&condition);
        rule.formulae.push(condition);
    }
    if let Some(style_name) = attr_value(start, b"table:apply-style-name") {
        rule.formulae.push(format!("apply-style:{}", style_name));
    }
    rule
}

fn parse_odf_condition_operator(condition: &str) -> Option<String> {
    let lower = condition.to_ascii_lowercase();
    if let Some(idx) = lower.find("cell-content-is-") {
        let rest = &lower[idx + "cell-content-is-".len()..];
        let op = rest.split('(').next().unwrap_or(rest);
        return Some(op.to_string());
    }
    if let Some(idx) = lower.find("is-true-formula") {
        let _ = idx;
        return Some("true-formula".to_string());
    }
    if let Some(idx) = lower.find("formula-is") {
        let _ = idx;
        return Some("formula".to_string());
    }
    None
}

#[derive(Debug, Clone)]
struct OdsRow {
    cells: Vec<OdsCellData>,
}

#[derive(Debug, Clone)]
struct OdsCellData {
    value: CellValue,
    formula: Option<CellFormula>,
    style_id: Option<u32>,
    col_repeat: u32,
    validation_name: Option<String>,
    col_span: Option<u32>,
    row_span: Option<u32>,
    is_covered: bool,
}

impl OdsCellData {
    fn should_emit(&self) -> bool {
        !matches!(self.value, CellValue::Empty) || self.formula.is_some()
    }

    fn merge_range(&self, row: u32, col: u32) -> Option<MergedCellRange> {
        let col_span = self.col_span.unwrap_or(1);
        let row_span = self.row_span.unwrap_or(1);
        if col_span > 1 || row_span > 1 {
            Some(MergedCellRange {
                start_col: col,
                start_row: row,
                end_col: col + col_span - 1,
                end_row: row + row_span - 1,
            })
        } else {
            None
        }
    }
}

fn parse_ods_row(
    reader: &mut OdfReader<'_>,
    _start: &BytesStart<'_>,
    store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<OdsRow, ParseError> {
    let mut buf = Vec::new();
    let mut cells = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table-cell" => {
                    let cell = parse_ods_cell(reader, &e, store, style_map, next_style_id)?;
                    cells.push(cell);
                }
                b"table:covered-table-cell" => {
                    let cell = parse_ods_covered_cell(reader, &e)?;
                    cells.push(cell);
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-cell" => {
                    let cell = parse_ods_cell_empty(&e, style_map, next_style_id)?;
                    cells.push(cell);
                }
                b"table:covered-table-cell" => {
                    let cell = parse_ods_covered_cell_empty(&e)?;
                    cells.push(cell);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table-row" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(OdsRow { cells })
}

fn parse_ods_row_sample(
    reader: &mut OdfReader<'_>,
    _start: &BytesStart<'_>,
    store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
    sample_cols: u32,
) -> Result<OdsRow, ParseError> {
    let mut buf = Vec::new();
    let mut cells = Vec::new();
    let mut col_idx: u32 = 0;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table-cell" => {
                    if col_idx >= sample_cols {
                        let repeat = attr_value(&e, b"table:number-columns-repeated")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        col_idx = col_idx.saturating_add(repeat);
                        spreadsheet::skip_element(reader, e.name().as_ref())?;
                    } else {
                        let cell = parse_ods_cell(reader, &e, store, style_map, next_style_id)?;
                        col_idx = col_idx.saturating_add(cell.col_repeat);
                        cells.push(cell);
                    }
                }
                b"table:covered-table-cell" => {
                    if col_idx >= sample_cols {
                        let repeat = attr_value(&e, b"table:number-columns-repeated")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        col_idx = col_idx.saturating_add(repeat);
                        spreadsheet::skip_element(reader, e.name().as_ref())?;
                    } else {
                        let cell = parse_ods_covered_cell(reader, &e)?;
                        col_idx = col_idx.saturating_add(cell.col_repeat);
                        cells.push(cell);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-cell" => {
                    if col_idx < sample_cols {
                        let cell = parse_ods_cell_empty(&e, style_map, next_style_id)?;
                        col_idx = col_idx.saturating_add(cell.col_repeat);
                        cells.push(cell);
                    } else {
                        let repeat = attr_value(&e, b"table:number-columns-repeated")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        col_idx = col_idx.saturating_add(repeat);
                    }
                }
                b"table:covered-table-cell" => {
                    if col_idx < sample_cols {
                        let cell = parse_ods_covered_cell_empty(&e)?;
                        col_idx = col_idx.saturating_add(cell.col_repeat);
                        cells.push(cell);
                    } else {
                        let repeat = attr_value(&e, b"table:number-columns-repeated")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        col_idx = col_idx.saturating_add(repeat);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table-row" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(OdsRow { cells })
}

#[derive(Debug, Clone)]
struct CellRef {
    sheet: Option<String>,
    row: u32,
    col: u32,
}

#[derive(Debug, Clone)]
struct CellRange {
    start: CellRef,
    end: CellRef,
}

#[derive(Debug, Clone)]
enum FormulaToken {
    Number(f64),
    Ident(String),
    Ref(CellRef),
    Range(CellRange),
    Plus,
    Minus,
    Star,
    Slash,
    LParen,
    RParen,
    Comma,
    End,
}

struct FormulaEvalContext<'a> {
    sheet_name: &'a str,
    values: HashMap<(u32, u32), CellValue>,
    formulas: &'a HashMap<(u32, u32), String>,
    cache: HashMap<(u32, u32), Option<f64>>,
    stack: Vec<(u32, u32)>,
}

impl<'a> FormulaEvalContext<'a> {
    fn new(
        sheet_name: &'a str,
        values: HashMap<(u32, u32), CellValue>,
        formulas: &'a HashMap<(u32, u32), String>,
    ) -> Self {
        Self {
            sheet_name,
            values,
            formulas,
            cache: HashMap::new(),
            stack: Vec::new(),
        }
    }

    fn eval_formula(&mut self, formula: &str) -> Option<f64> {
        let tokens = tokenize_formula(formula);
        let mut parser = FormulaParser::new(tokens, self);
        parser.parse_expression()
    }

    fn resolve_ref(&mut self, reference: &CellRef) -> Option<f64> {
        if let Some(sheet) = reference.sheet.as_deref() {
            if !sheet.eq_ignore_ascii_case(self.sheet_name) {
                return None;
            }
        }
        let key = (reference.row, reference.col);
        if let Some(value) = self.cache.get(&key) {
            return *value;
        }
        if self.stack.contains(&key) {
            self.cache.insert(key, None);
            return None;
        }
        if let Some(value) = self.values.get(&key) {
            if let Some(number) = cell_value_to_number(value) {
                self.cache.insert(key, Some(number));
                return Some(number);
            }
        }
        let formula_text = self.formulas.get(&key)?.clone();
        self.stack.push(key);
        let result = self.eval_formula(&formula_text);
        self.stack.pop();
        if let Some(number) = result {
            self.values.insert(key, CellValue::Number(number));
        }
        self.cache.insert(key, result);
        result
    }

    fn resolve_range(&mut self, range: &CellRange) -> Option<Vec<f64>> {
        if let Some(sheet) = range.start.sheet.as_deref() {
            if !sheet.eq_ignore_ascii_case(self.sheet_name) {
                return None;
            }
        }
        if let Some(sheet) = range.end.sheet.as_deref() {
            if !sheet.eq_ignore_ascii_case(self.sheet_name) {
                return None;
            }
        }
        let row_start = range.start.row.min(range.end.row);
        let row_end = range.start.row.max(range.end.row);
        let col_start = range.start.col.min(range.end.col);
        let col_end = range.start.col.max(range.end.col);
        let total = (row_end - row_start + 1) as u64 * (col_end - col_start + 1) as u64;
        if total > 1_000_000 {
            return None;
        }
        let mut values = Vec::new();
        for row in row_start..=row_end {
            for col in col_start..=col_end {
                let reference = CellRef {
                    sheet: None,
                    row,
                    col,
                };
                if let Some(number) = self.resolve_ref(&reference) {
                    values.push(number);
                }
            }
        }
        Some(values)
    }
}

struct FormulaParser<'a, 'b> {
    tokens: Vec<FormulaToken>,
    pos: usize,
    ctx: &'a mut FormulaEvalContext<'b>,
}

impl<'a, 'b> FormulaParser<'a, 'b> {
    fn new(tokens: Vec<FormulaToken>, ctx: &'a mut FormulaEvalContext<'b>) -> Self {
        Self {
            tokens,
            pos: 0,
            ctx,
        }
    }

    fn parse_expression(&mut self) -> Option<f64> {
        let mut value = self.parse_term()?;
        loop {
            match self.peek() {
                FormulaToken::Plus => {
                    self.next();
                    value += self.parse_term()?;
                }
                FormulaToken::Minus => {
                    self.next();
                    value -= self.parse_term()?;
                }
                _ => break,
            }
        }
        Some(value)
    }

    fn parse_term(&mut self) -> Option<f64> {
        let mut value = self.parse_factor()?;
        loop {
            match self.peek() {
                FormulaToken::Star => {
                    self.next();
                    value *= self.parse_factor()?;
                }
                FormulaToken::Slash => {
                    self.next();
                    let denom = self.parse_factor()?;
                    if denom == 0.0 {
                        return None;
                    }
                    value /= denom;
                }
                _ => break,
            }
        }
        Some(value)
    }

    fn parse_factor(&mut self) -> Option<f64> {
        let token = self.peek().clone();
        match token {
            FormulaToken::Minus => {
                self.next();
                self.parse_factor().map(|v| -v)
            }
            FormulaToken::Number(value) => {
                self.next();
                Some(value)
            }
            FormulaToken::Ref(reference) => {
                self.next();
                self.ctx.resolve_ref(&reference)
            }
            FormulaToken::Range(range) => {
                self.next();
                let values = self.ctx.resolve_range(&range)?;
                Some(values.iter().sum())
            }
            FormulaToken::Ident(name) => {
                self.next();
                if matches!(self.peek(), FormulaToken::LParen) {
                    self.next();
                    let values = self.parse_function_args()?;
                    if !matches!(self.peek(), FormulaToken::RParen) {
                        return None;
                    }
                    self.next();
                    eval_formula_function(&name, &values)
                } else {
                    None
                }
            }
            FormulaToken::LParen => {
                self.next();
                let value = self.parse_expression()?;
                if !matches!(self.peek(), FormulaToken::RParen) {
                    return None;
                }
                self.next();
                Some(value)
            }
            _ => None,
        }
    }

    fn parse_function_args(&mut self) -> Option<Vec<f64>> {
        let mut values = Vec::new();
        if matches!(self.peek(), FormulaToken::RParen) {
            return Some(values);
        }
        loop {
            if matches!(self.peek(), FormulaToken::Range(_)) {
                if let FormulaToken::Range(range) = self.next().clone() {
                    let range_values = self.ctx.resolve_range(&range)?;
                    values.extend(range_values);
                }
            } else {
                let value = self.parse_expression()?;
                values.push(value);
            }
            match self.peek() {
                FormulaToken::Comma => {
                    self.next();
                }
                FormulaToken::RParen => break,
                _ => return None,
            }
        }
        Some(values)
    }

    fn peek(&self) -> &FormulaToken {
        self.tokens.get(self.pos).unwrap_or(&FormulaToken::End)
    }

    fn next(&mut self) -> &FormulaToken {
        let token = self.tokens.get(self.pos).unwrap_or(&FormulaToken::End);
        self.pos += 1;
        token
    }
}

fn tokenize_formula(formula: &str) -> Vec<FormulaToken> {
    let mut tokens = Vec::new();
    let mut chars = formula.trim().chars().peekable();
    while let Some(&ch) = chars.peek() {
        match ch {
            ' ' | '\t' | '\n' | '\r' => {
                chars.next();
            }
            '+' => {
                chars.next();
                tokens.push(FormulaToken::Plus);
            }
            '-' => {
                chars.next();
                tokens.push(FormulaToken::Minus);
            }
            '*' => {
                chars.next();
                tokens.push(FormulaToken::Star);
            }
            '/' => {
                chars.next();
                tokens.push(FormulaToken::Slash);
            }
            '(' => {
                chars.next();
                tokens.push(FormulaToken::LParen);
            }
            ')' => {
                chars.next();
                tokens.push(FormulaToken::RParen);
            }
            ',' | ';' => {
                chars.next();
                tokens.push(FormulaToken::Comma);
            }
            '[' => {
                chars.next();
                let mut buffer = String::new();
                while let Some(c) = chars.next() {
                    if c == ']' {
                        break;
                    }
                    buffer.push(c);
                }
                if let Some(token) = parse_bracket_reference(&buffer) {
                    tokens.push(token);
                }
            }
            _ => {
                if ch.is_ascii_digit() || ch == '.' {
                    let mut num = String::new();
                    while let Some(&c) = chars.peek() {
                        if c.is_ascii_digit() || c == '.' {
                            num.push(c);
                            chars.next();
                        } else {
                            break;
                        }
                    }
                    if let Ok(value) = num.parse::<f64>() {
                        tokens.push(FormulaToken::Number(value));
                    }
                } else if ch.is_ascii_alphabetic() || ch == '_' {
                    let mut ident = String::new();
                    while let Some(&c) = chars.peek() {
                        if c.is_ascii_alphanumeric() || c == '_' || c == '.' || c == '$' {
                            ident.push(c);
                            chars.next();
                        } else {
                            break;
                        }
                    }
                    if let Some(reference) = parse_simple_reference(&ident) {
                        tokens.push(FormulaToken::Ref(reference));
                    } else {
                        tokens.push(FormulaToken::Ident(ident));
                    }
                } else {
                    chars.next();
                }
            }
        }
    }
    tokens.push(FormulaToken::End);
    tokens
}

fn parse_bracket_reference(input: &str) -> Option<FormulaToken> {
    let trimmed = input.trim();
    if let Some((start, end)) = trimmed.split_once(':') {
        let start_ref = parse_sheeted_cell(start)?;
        let end_ref = parse_sheeted_cell(end)?;
        return Some(FormulaToken::Range(CellRange {
            start: start_ref,
            end: end_ref,
        }));
    }
    parse_sheeted_cell(trimmed).map(FormulaToken::Ref)
}

fn parse_simple_reference(input: &str) -> Option<CellRef> {
    if input.chars().any(|c| c.is_ascii_digit()) && input.chars().any(|c| c.is_ascii_alphabetic()) {
        parse_sheeted_cell(input)
    } else {
        None
    }
}

fn parse_sheeted_cell(input: &str) -> Option<CellRef> {
    let trimmed = input.trim().trim_start_matches('.');
    let mut sheet: Option<String> = None;
    let mut cell_part = trimmed;
    if let Some((sheet_part, cell)) = trimmed.rsplit_once('.') {
        if !sheet_part.is_empty() && !cell.is_empty() {
            let sheet_name = sheet_part.trim_matches('\'').replace('$', "");
            if !sheet_name.is_empty() {
                sheet = Some(sheet_name);
            }
            cell_part = cell;
        }
    }
    let cell = parse_cell_ref(cell_part)?;
    Some(CellRef {
        sheet,
        row: cell.row,
        col: cell.col,
    })
}

fn parse_cell_ref(input: &str) -> Option<CellRef> {
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in input.chars() {
        if ch.is_ascii_alphabetic() {
            letters.push(ch);
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else if ch == '$' {
            continue;
        } else {
            break;
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    let col = column_name_to_index(&letters)?;
    let row = digits.parse::<u32>().ok()?.saturating_sub(1);
    Some(CellRef {
        sheet: None,
        row,
        col,
    })
}

fn column_name_to_index(name: &str) -> Option<u32> {
    let mut index: u32 = 0;
    for ch in name.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        index = index * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    Some(index.saturating_sub(1))
}

fn cell_value_to_number(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Number(num) => Some(*num),
        CellValue::Boolean(v) => Some(if *v { 1.0 } else { 0.0 }),
        CellValue::String(s) => s.parse::<f64>().ok(),
        _ => None,
    }
}

fn eval_formula_function(name: &str, values: &[f64]) -> Option<f64> {
    let upper = name.to_ascii_uppercase();
    match upper.as_str() {
        "SUM" => Some(values.iter().sum()),
        "AVERAGE" => {
            if values.is_empty() {
                None
            } else {
                Some(values.iter().sum::<f64>() / values.len() as f64)
            }
        }
        "MIN" => values.iter().copied().reduce(f64::min),
        "MAX" => values.iter().copied().reduce(f64::max),
        "COUNT" => Some(values.len() as f64),
        _ => None,
    }
}

fn evaluate_ods_formulas(
    sheet_name: &str,
    formula_cells: &[(NodeId, u32, u32, String)],
    store: &mut IrStore,
    cell_values: &mut HashMap<(u32, u32), CellValue>,
    formula_map: &HashMap<(u32, u32), String>,
) {
    let mut ctx = FormulaEvalContext::new(sheet_name, cell_values.clone(), formula_map);
    for (cell_id, row, col, formula) in formula_cells {
        if let Some(IRNode::Cell(cell)) = store.get_mut(*cell_id) {
            if matches!(cell.value, CellValue::Empty) {
                if let Some(value) = ctx.eval_formula(formula) {
                    cell.value = CellValue::Number(value);
                    cell_values.insert((*row, *col), CellValue::Number(value));
                }
            }
        }
    }
}

fn parse_ods_covered_cell(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
) -> Result<OdsCellData, ParseError> {
    let col_repeat = attr_value(start, b"table:number-columns-repeated")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let col_span =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok());
    let row_span =
        attr_value(start, b"table:number-rows-spanned").and_then(|v| v.parse::<u32>().ok());
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::End(e)) if e.name().as_ref() == b"table:covered-table-cell" => break,
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(OdsCellData {
        value: CellValue::Empty,
        formula: None,
        style_id: None,
        col_repeat,
        validation_name: None,
        col_span,
        row_span,
        is_covered: true,
    })
}

fn parse_ods_covered_cell_empty(start: &BytesStart<'_>) -> Result<OdsCellData, ParseError> {
    let col_repeat = attr_value(start, b"table:number-columns-repeated")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let col_span =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok());
    let row_span =
        attr_value(start, b"table:number-rows-spanned").and_then(|v| v.parse::<u32>().ok());
    Ok(OdsCellData {
        value: CellValue::Empty,
        formula: None,
        style_id: None,
        col_repeat,
        validation_name: None,
        col_span,
        row_span,
        is_covered: true,
    })
}

fn parse_text_element(reader: &mut OdfReader<'_>, end_name: &[u8]) -> Result<String, ParseError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:s" => {
                    let count = attr_value(&e, b"text:c")
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1);
                    text.extend(std::iter::repeat(' ').take(count));
                }
                b"text:tab" => text.push('\t'),
                b"text:line-break" => text.push('\n'),
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"text:s" => {
                    let count = attr_value(&e, b"text:c")
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1);
                    text.extend(std::iter::repeat(' ').take(count));
                }
                b"text:tab" => text.push('\t'),
                b"text:line-break" => text.push('\n'),
                _ => {}
            },
            Ok(Event::Text(e)) => {
                let chunk = e.unescape().unwrap_or_default();
                text.push_str(&chunk);
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_name {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

fn parse_draw_frame_spreadsheet(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let transform = parse_frame_transform(start);
    let mut shape_type = ShapeType::Picture;
    let mut media_target: Option<String> = None;
    let mut buf = Vec::new();
    let mut has_shape = false;
    let mut name = attr_value(start, b"draw:name");
    let mut chart_id: Option<NodeId> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"draw:image" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href);
                        shape_type = ShapeType::Picture;
                        has_shape = true;
                    }
                }
                b"chart:chart" => {
                    shape_type = ShapeType::Chart;
                    has_shape = true;
                    let chart = parse_odf_chart(reader, &e)?;
                    let id = chart.id;
                    store.insert(IRNode::ChartData(chart));
                    chart_id = Some(id);
                }
                b"draw:plugin" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                        shape_type = classify_media_shape(&href);
                        has_shape = true;
                    }
                }
                b"draw:object" | b"draw:object-ole" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                    }
                    shape_type = ShapeType::OleObject;
                    has_shape = true;
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"draw:image" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href);
                        shape_type = ShapeType::Picture;
                        has_shape = true;
                    }
                }
                b"chart:chart" => {
                    shape_type = ShapeType::Chart;
                    has_shape = true;
                    let mut chart = ChartData::new();
                    chart.chart_type = attr_value(&e, b"chart:class");
                    chart.span = Some(SourceSpan::new("content.xml"));
                    let id = chart.id;
                    store.insert(IRNode::ChartData(chart));
                    chart_id = Some(id);
                }
                b"draw:plugin" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                        shape_type = classify_media_shape(&href);
                        has_shape = true;
                    }
                }
                b"draw:object" | b"draw:object-ole" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                    }
                    shape_type = ShapeType::OleObject;
                    has_shape = true;
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"draw:frame" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if has_shape {
        let mut shape = Shape::new(shape_type);
        shape.name = name.take();
        shape.media_target = media_target;
        shape.chart_id = chart_id;
        shape.transform = transform;
        let shape_id = shape.id;
        store.insert(IRNode::Shape(shape));
        Ok(Some(shape_id))
    } else {
        Ok(None)
    }
}

fn column_index_to_name(mut index: u32) -> String {
    let mut name = String::new();
    index += 1;
    while index > 0 {
        let rem = ((index - 1) % 26) as u8;
        name.push((b'A' + rem) as char);
        index = (index - 1) / 26;
    }
    name.chars().rev().collect()
}

#[derive(Debug, Clone, Copy)]
struct ListContext {
    num_id: u32,
    level: u32,
}

fn parse_table(
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut buf = Vec::new();
    let mut table = Table::new();
    let mut current_row: Option<TableRow> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    current_row = Some(TableRow::new());
                }
                b"table:table-cell" => {
                    let cell_id = parse_table_cell(reader, &e, store, limits)?;
                    if let Some(row) = current_row.as_mut() {
                        row.cells.push(cell_id);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    let row = TableRow::new();
                    let row_id = row.id;
                    store.insert(IRNode::TableRow(row));
                    table.rows.push(row_id);
                }
                b"table:table-cell" => {
                    let mut cell = TableCell::new();
                    if let Some(span) = attr_value(&e, b"table:number-columns-spanned")
                        .and_then(|v| v.parse::<u32>().ok())
                    {
                        let mut props = TableCellProperties::default();
                        props.grid_span = Some(span);
                        cell.properties = props;
                    }
                    let cell_id = cell.id;
                    store.insert(IRNode::TableCell(cell));
                    if let Some(row) = current_row.as_mut() {
                        row.cells.push(cell_id);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    if let Some(row) = current_row.take() {
                        let row_id = row.id;
                        store.insert(IRNode::TableRow(row));
                        table.rows.push(row_id);
                    }
                }
                b"table:table" => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    let table_id = table.id;
    store.insert(IRNode::Table(table));
    Ok(table_id)
}

fn parse_table_cell(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut cell = TableCell::new();
    if let Some(span) =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok())
    {
        let mut props = TableCellProperties::default();
        props.grid_span = Some(span);
        cell.properties = props;
    }

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:p" => {
                    let paragraph_id = parse_paragraph(
                        reader,
                        e.name().as_ref(),
                        None,
                        None,
                        store,
                        &mut Vec::new(),
                        limits,
                    )?;
                    cell.content.push(paragraph_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table-cell" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    let cell_id = cell.id;
    store.insert(IRNode::TableCell(cell));
    Ok(cell_id)
}

fn parse_annotation(
    reader: &mut OdfReader<'_>,
    comment_id: &str,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut comment = Comment::new(comment_id);
    let mut buf = Vec::new();
    let mut current = None;

    #[derive(Clone, Copy)]
    enum AnnotationField {
        Creator,
        Date,
    }

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"dc:creator" => current = Some(AnnotationField::Creator),
                b"dc:date" => current = Some(AnnotationField::Date),
                b"text:p" => {
                    let paragraph_id = parse_paragraph(
                        reader,
                        e.name().as_ref(),
                        None,
                        None,
                        store,
                        &mut Vec::new(),
                        limits,
                    )?;
                    comment.content.push(paragraph_id);
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if let Some(field) = current {
                    let value = e.unescape().unwrap_or_default().to_string();
                    match field {
                        AnnotationField::Creator => comment.author = Some(value),
                        AnnotationField::Date => comment.date = Some(value),
                    }
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"office:annotation" => break,
                b"dc:creator" | b"dc:date" => current = None,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    let comment_id = comment.id;
    store.insert(IRNode::Comment(comment));
    Ok(comment_id)
}

fn parse_note(
    reader: &mut OdfReader<'_>,
    note_id: &str,
    note_class: &str,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut buf = Vec::new();
    let mut content: Vec<NodeId> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:p" => {
                    let paragraph_id = parse_paragraph(
                        reader,
                        e.name().as_ref(),
                        None,
                        None,
                        store,
                        &mut Vec::new(),
                        limits,
                    )?;
                    content.push(paragraph_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"text:note" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if note_class == "endnote" {
        let mut endnote = Endnote::new(note_id);
        endnote.content = content;
        let id = endnote.id;
        store.insert(IRNode::Endnote(endnote));
        Ok(id)
    } else {
        let mut footnote = Footnote::new(note_id);
        footnote.content = content;
        let id = footnote.id;
        store.insert(IRNode::Footnote(footnote));
        Ok(id)
    }
}

fn parse_draw_frame(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let mut shape = Shape::new(ShapeType::Picture);
    shape.transform = parse_frame_transform(start);
    shape.name = attr_value(start, b"draw:name");
    let mut buf = Vec::new();
    let mut has_shape = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"draw:image" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        shape.media_target = Some(href);
                        shape.shape_type = ShapeType::Picture;
                        has_shape = true;
                    }
                }
                b"draw:object" | b"draw:object-ole" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        shape.media_target = Some(href);
                    }
                    shape.shape_type = ShapeType::OleObject;
                    has_shape = true;
                }
                b"draw:plugin" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        shape.media_target = Some(href.clone());
                        shape.shape_type = classify_media_shape(&href);
                        has_shape = true;
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"draw:frame" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if has_shape {
        let shape_id = shape.id;
        store.insert(IRNode::Shape(shape));
        Ok(Some(shape_id))
    } else {
        Ok(None)
    }
}

fn parse_tracked_changes(
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<Vec<NodeId>, ParseError> {
    let mut buf = Vec::new();
    let mut revisions = Vec::new();
    let mut current_revision: Option<Revision> = None;
    let mut current_field: Option<ChangeInfoField> = None;

    enum ChangeInfoField {
        Author,
        Date,
    }

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:changed-region" => {
                    current_revision = None;
                }
                b"text:insertion" => {
                    current_revision = Some(Revision::new(RevisionType::Insert));
                }
                b"text:deletion" => {
                    current_revision = Some(Revision::new(RevisionType::Delete));
                }
                b"dc:creator" => current_field = Some(ChangeInfoField::Author),
                b"dc:date" => current_field = Some(ChangeInfoField::Date),
                b"text:p" => {
                    if let Some(rev) = current_revision.as_mut() {
                        let paragraph_id = parse_paragraph(
                            reader,
                            e.name().as_ref(),
                            None,
                            None,
                            store,
                            &mut Vec::new(),
                            limits,
                        )?;
                        rev.content.push(paragraph_id);
                    }
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if let Some(rev) = current_revision.as_mut() {
                    if let Some(field) = &current_field {
                        let value = e.unescape().unwrap_or_default().to_string();
                        match field {
                            ChangeInfoField::Author => rev.author = Some(value),
                            ChangeInfoField::Date => rev.date = Some(value),
                        }
                    }
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"text:insertion" | b"text:deletion" => {
                    if let Some(rev) = current_revision.take() {
                        let id = rev.id;
                        store.insert(IRNode::Revision(rev));
                        revisions.push(id);
                    }
                }
                b"text:tracked-changes" => break,
                b"dc:creator" | b"dc:date" => current_field = None,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(revisions)
}

fn parse_styles(xml: &str) -> Option<StyleSet> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut styles = StyleSet::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"style:style" => {
                    if let Some(mut style) = build_style_from_start(&e, false) {
                        parse_style_properties(&mut reader, &mut style, b"style:style");
                        styles.styles.push(style);
                    }
                }
                b"style:default-style" => {
                    if let Some(mut style) = build_style_from_start(&e, true) {
                        parse_style_properties(&mut reader, &mut style, b"style:default-style");
                        styles.styles.push(style);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"style:style" => {
                    if let Some(style) = build_style_from_start(&e, false) {
                        styles.styles.push(style);
                    }
                }
                b"style:default-style" => {
                    if let Some(style) = build_style_from_start(&e, true) {
                        styles.styles.push(style);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }

    if styles.styles.is_empty() {
        None
    } else {
        Some(styles)
    }
}

fn map_style_family(e: &BytesStart<'_>) -> StyleType {
    match attr_value(e, b"style:family").as_deref() {
        Some("paragraph") => StyleType::Paragraph,
        Some("text") => StyleType::Character,
        Some("table") => StyleType::Table,
        Some("list") => StyleType::Numbering,
        _ => StyleType::Other,
    }
}

fn build_style_from_start(start: &BytesStart<'_>, is_default: bool) -> Option<Style> {
    let style_id = attr_value(start, b"style:name")
        .or_else(|| attr_value(start, b"style:family").map(|f| format!("default:{f}")));
    let style_id = style_id?;
    let mut style = Style {
        style_id,
        name: attr_value(start, b"style:display-name"),
        style_type: map_style_family(start),
        based_on: attr_value(start, b"style:parent-style-name"),
        next: attr_value(start, b"style:next-style-name"),
        is_default,
        run_props: None,
        paragraph_props: None,
        table_props: None,
    };
    if let Some(family) = attr_value(start, b"style:family") {
        if family == "paragraph" || family == "text" {
            style.is_default = is_default
                || attr_value(start, b"style:default")
                    .map(|v| v == "true")
                    .unwrap_or(false);
        }
    }
    Some(style)
}

fn parse_style_properties(reader: &mut Reader<&[u8]>, style: &mut Style, end_name: &[u8]) {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"style:text-properties" => {
                    let mut props = style.run_props.take().unwrap_or_default();
                    if let Some(font) = attr_value(&e, b"fo:font-family")
                        .or_else(|| attr_value(&e, b"style:font-name"))
                    {
                        props.font_family = Some(font);
                    }
                    if let Some(size) =
                        attr_value(&e, b"fo:font-size").and_then(|v| parse_font_size(&v))
                    {
                        props.font_size = Some(size);
                    }
                    if let Some(weight) = attr_value(&e, b"fo:font-weight") {
                        props.bold = Some(weight.eq_ignore_ascii_case("bold"));
                    }
                    if let Some(style_attr) = attr_value(&e, b"fo:font-style") {
                        props.italic = Some(style_attr.eq_ignore_ascii_case("italic"));
                    }
                    if let Some(color) = attr_value(&e, b"fo:color") {
                        props.color = Some(color);
                    }
                    style.run_props = Some(props);
                }
                b"style:paragraph-properties" => {
                    let mut props = style.paragraph_props.take().unwrap_or_default();
                    if let Some(align) =
                        attr_value(&e, b"fo:text-align").and_then(|v| parse_text_alignment(&v))
                    {
                        props.alignment = Some(align);
                    }
                    style.paragraph_props = Some(props);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_name {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

fn parse_font_size(value: &str) -> Option<u32> {
    let trimmed = value.trim();
    let num = trimmed
        .trim_end_matches("pt")
        .trim_end_matches("px")
        .trim_end_matches("cm")
        .trim_end_matches("mm");
    num.parse::<f32>().ok().map(|v| v.round() as u32)
}

fn merge_styles(existing: &mut StyleSet, incoming: &mut StyleSet) {
    let mut seen = existing
        .styles
        .iter()
        .map(|s| s.style_id.clone())
        .collect::<std::collections::HashSet<String>>();
    for style in incoming.styles.drain(..) {
        if seen.insert(style.style_id.clone()) {
            existing.styles.push(style);
        }
    }
}

fn parse_master_pages(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"style:master-page" {
                    if let Some(name) = attr_value(&e, b"style:name") {
                        out.push(name);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn parse_page_layouts(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"style:page-layout" {
                    if let Some(name) = attr_value(&e, b"style:name") {
                        out.push(name);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn parse_odf_headers_footers(
    xml: &str,
    store: &mut IrStore,
    config: &ParserConfig,
) -> Result<(Vec<NodeId>, Vec<NodeId>), ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml.as_bytes()));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut headers = Vec::new();
    let mut footers = Vec::new();

    let limits = OdfLimits::new(config, false);

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"style:header" | b"style:header-left" => {
                    let content = parse_odf_header_footer_block(
                        &mut reader,
                        e.name().as_ref(),
                        store,
                        &limits,
                    )?;
                    let mut header = Header::new();
                    header.content = content;
                    header.span = Some(SourceSpan::new("styles.xml"));
                    let id = header.id;
                    store.insert(IRNode::Header(header));
                    headers.push(id);
                }
                b"style:footer" | b"style:footer-left" => {
                    let content = parse_odf_header_footer_block(
                        &mut reader,
                        e.name().as_ref(),
                        store,
                        &limits,
                    )?;
                    let mut footer = Footer::new();
                    footer.content = content;
                    footer.span = Some(SourceSpan::new("styles.xml"));
                    let id = footer.id;
                    store.insert(IRNode::Footer(footer));
                    footers.push(id);
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "styles.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok((headers, footers))
}

fn parse_odf_header_footer_block(
    reader: &mut OdfReader<'_>,
    end_name: &[u8],
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<Vec<NodeId>, ParseError> {
    let mut buf = Vec::new();
    let mut content = Vec::new();
    let mut list_stack: Vec<ListContext> = Vec::new();
    let mut list_id_map: HashMap<String, u32> = HashMap::new();
    let mut next_list_id = 1u32;
    let mut pending_inline_nodes: Vec<NodeId> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:list" => {
                    let style_name = attr_value(&e, b"text:style-name").unwrap_or_default();
                    let num_id = list_id_map.entry(style_name).or_insert_with(|| {
                        let id = next_list_id;
                        next_list_id += 1;
                        id
                    });
                    let level = list_stack.len() as u32;
                    list_stack.push(ListContext {
                        num_id: *num_id,
                        level,
                    });
                }
                b"text:p" | b"text:h" => {
                    let outline_level =
                        attr_value(&e, b"text:outline-level").and_then(|v| v.parse::<u8>().ok());
                    let numbering = list_stack.last().map(|ctx| NumberingInfo {
                        num_id: ctx.num_id,
                        level: ctx.level,
                        format: None,
                    });
                    let paragraph_id = parse_paragraph(
                        reader,
                        e.name().as_ref(),
                        numbering,
                        outline_level,
                        store,
                        &mut pending_inline_nodes,
                        limits,
                    )?;
                    content.extend(pending_inline_nodes.drain(..));
                    content.push(paragraph_id);
                }
                b"table:table" => {
                    let table_id = parse_table(reader, store, limits)?;
                    content.extend(pending_inline_nodes.drain(..));
                    content.push(table_id);
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"text:list" => {}
                b"text:p" | b"text:h" => {
                    let outline_level =
                        attr_value(&e, b"text:outline-level").and_then(|v| v.parse::<u8>().ok());
                    let numbering = list_stack.last().map(|ctx| NumberingInfo {
                        num_id: ctx.num_id,
                        level: ctx.level,
                        format: None,
                    });
                    let paragraph_id = text::build_paragraph(store, "", numbering, outline_level);
                    content.extend(pending_inline_nodes.drain(..));
                    content.push(paragraph_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"text:list" => {
                    list_stack.pop();
                }
                _ if e.name().as_ref() == end_name => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "styles.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(content)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::DocumentParser;
    use docir_core::security::ThreatIndicatorType;
    use std::io::{Cursor, Write};
    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_odf_zip(mimetype: &str, content_xml: &str, styles_xml: Option<&str>) -> Vec<u8> {
        build_odf_zip_custom(mimetype, content_xml, styles_xml, None, Vec::new())
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
        let stored =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
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
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
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
"#;
        let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Scripts/macro.py" manifest:media-type="application/vnd.sun.star.script"/>
  <manifest:file-entry manifest:full-path="Object 1" manifest:media-type="application/vnd.sun.star.oleobject"/>
  <manifest:file-entry manifest:full-path="Encrypted" manifest:media-type="application/vnd.oasis.opendocument.encrypted"/>
</manifest:manifest>
"#;
        let signatures_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
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
"#;
        let zip_data = build_odf_zip_custom(
            mimetype,
            content_xml,
            None,
            Some(manifest_xml),
            vec![
                ("META-INF/documentsignatures.xml", signatures_xml),
                ("Scripts/macro.py", b"print('hi')"),
                ("Object 1", b"oledata"),
            ],
        );
        let parser = DocumentParser::new();
        let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
        docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
        let doc = parsed.document().unwrap();

        assert!(doc.security.macro_project.is_some());
        assert!(!doc.security.external_refs.is_empty());
        assert!(!doc.security.ole_objects.is_empty());

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

        let mut table_count = 0;
        let mut comment_count = 0;
        let mut footnote_count = 0;
        let mut bookmark_count = 0;
        let mut field_count = 0;
        let mut shape_count = 0;
        let mut revision_count = 0;
        let mut styles_count = 0;

        for node in parsed.store.values() {
            match node {
                IRNode::Table(_) => table_count += 1,
                IRNode::Comment(_) => comment_count += 1,
                IRNode::Footnote(_) => footnote_count += 1,
                IRNode::BookmarkStart(_) | IRNode::BookmarkEnd(_) => bookmark_count += 1,
                IRNode::Field(_) => field_count += 1,
                IRNode::Shape(_) => shape_count += 1,
                IRNode::Revision(_) => revision_count += 1,
                IRNode::StyleSet(_) => styles_count += 1,
                _ => {}
            }
        }

        assert_eq!(table_count, 1);
        assert_eq!(comment_count, 1);
        assert_eq!(footnote_count, 1);
        assert!(bookmark_count >= 2);
        assert_eq!(field_count, 1);
        assert_eq!(shape_count, 1);
        assert_eq!(revision_count, 1);
        assert_eq!(styles_count, 1);
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

    #[test]
    fn test_odf_security_threat_indicators() {
        let mimetype = "application/vnd.oasis.opendocument.text";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
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
"#;
        let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Scripts/macro.py" manifest:media-type="application/vnd.sun.star.script"/>
  <manifest:file-entry manifest:full-path="Object 1" manifest:media-type="application/vnd.sun.star.oleobject"/>
</manifest:manifest>
"#;
        let zip_data = build_odf_zip_custom(
            mimetype,
            content_xml,
            None,
            Some(manifest_xml),
            vec![
                ("Scripts/macro.py", b"print('hi')"),
                ("Object 1", b"oledata"),
            ],
        );
        let parser = DocumentParser::new();
        let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
        docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
        let doc = parsed.document().unwrap();

        assert!(doc.security.macro_project.is_some());
        assert!(!doc.security.external_refs.is_empty());
        assert!(!doc.security.ole_objects.is_empty());
        assert!(doc
            .security
            .threat_indicators
            .iter()
            .any(|i| i.indicator_type == ThreatIndicatorType::RemoteResource));
        assert!(doc
            .security
            .threat_indicators
            .iter()
            .any(|i| i.indicator_type == ThreatIndicatorType::OleObject));
    }
}
