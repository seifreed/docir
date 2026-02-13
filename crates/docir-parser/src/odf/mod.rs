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
    ConditionalFormat, ConditionalRule, DataValidation, DiagnosticSeverity, Diagnostics, Document,
    Endnote, ExtensionPart, ExtensionPartKind, Field, FieldInstruction, FieldKind, Footer,
    Footnote, Header, IRNode, MediaAsset, MediaType, MergedCellRange, NumberingInfo, Paragraph,
    ParagraphProperties, PivotCache, PivotCacheRecords, PivotTable, Revision, RevisionType, Run,
    Section, Shape, ShapeText, ShapeTextParagraph, ShapeTextRun, ShapeType, Slide, SlideAnimation,
    SlideTransition, Style, StyleSet, StyleType, Table, TableCell, TableCellProperties, TableRow,
    Worksheet, WorksheetDrawing,
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
mod sampling;
pub(crate) mod security;
mod security_helpers;
mod spreadsheet;
mod text;
mod utils;

use self::security::scan_odf_filters;
use self::security_helpers::{build_odf_macro_project, parse_odf_signatures};
use crate::security_scan::{DefaultSecurityScanner, SecurityScanner};

use container::{handle_content_xml, load_meta};
use formula::evaluate_ods_formulas;
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
use sampling::parse_ods_row_sample;
use utils::{parse_frame_transform, parse_ods_named_ranges, strip_odf_formula_prefix};

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
mod tests;
