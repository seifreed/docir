//! ODF spreadsheet parsing helpers.

use super::super::helpers::{parse_validation_definition, ValidationDef};
#[path = "spreadsheet_frame.rs"]
mod spreadsheet_frame;
#[path = "spreadsheet_pivot.rs"]
mod spreadsheet_pivot;
use super::super::spreadsheet_chunks::{
    extract_spreadsheet_table_chunks, table_name_from_chunk, OdfTableChunk,
};
use super::super::{
    attr_value, parse_ods_table, parse_ods_table_fast, scan_xml_events,
    scan_xml_events_with_reader, OdfAtomicLimits, OdfContentResult, OdfLimitCounter, OdfReader,
};
use crate::error::ParseError;
use crate::parser::ParserConfig;
use crate::xml_utils::{is_end_event, scan_xml_events_until_end, XmlScanControl};
use docir_core::ir::{IRNode, Worksheet};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
pub(crate) use spreadsheet_frame::parse_draw_frame_spreadsheet;
use spreadsheet_pivot::{
    collect_validation_definitions, parse_ods_pivot_table_empty, parse_ods_pivot_table_full,
    parse_ods_pivots_from_xml, record_pivot_parse, PivotLinks,
};
use std::collections::HashMap;
use std::sync::Arc;
use std::thread;

struct FastSpreadsheetState<'a> {
    in_spreadsheet: &'a mut bool,
    sheet_id: &'a mut u32,
    store: &'a mut IrStore,
    sheets: &'a mut Vec<NodeId>,
    validations: &'a HashMap<String, ValidationDef>,
    limits: &'a dyn OdfLimitCounter,
}

struct FullSpreadsheetState<'a> {
    in_spreadsheet: &'a mut bool,
    sheet_id: &'a mut u32,
    store: &'a mut IrStore,
    limits: &'a dyn OdfLimitCounter,
    validations: &'a mut HashMap<String, ValidationDef>,
    sheets: &'a mut Vec<NodeId>,
    sheet_index: &'a mut HashMap<String, NodeId>,
    pivot_links: &'a mut Vec<(Option<String>, NodeId)>,
    pivot_caches: &'a mut Vec<NodeId>,
    next_cache_id: &'a mut u32,
}

pub(crate) fn parse_content_spreadsheet(
    xml: &[u8],
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<OdfContentResult, ParseError> {
    if limits.fast_mode() {
        return parse_content_spreadsheet_fast(xml, store, limits);
    }
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut in_spreadsheet = false;
    let mut sheet_id = 1u32;
    let mut sheets = Vec::new();
    let mut validations: HashMap<String, ValidationDef> = HashMap::new();
    let mut sheet_index: HashMap<String, NodeId> = HashMap::new();
    let mut pivot_links: PivotLinks = Vec::new();
    let mut pivot_caches: Vec<NodeId> = Vec::new();
    let mut next_cache_id: u32 = 1;

    scan_xml_events_until_end(
        &mut reader,
        &mut buf,
        "content.xml",
        |event| is_end_event(event, b"office:spreadsheet"),
        |reader, event| {
            match event {
                Event::Start(start) => {
                    handle_spreadsheet_start_full(
                        reader,
                        start,
                        FullSpreadsheetState {
                            in_spreadsheet: &mut in_spreadsheet,
                            sheet_id: &mut sheet_id,
                            store,
                            limits,
                            validations: &mut validations,
                            sheets: &mut sheets,
                            sheet_index: &mut sheet_index,
                            pivot_links: &mut pivot_links,
                            pivot_caches: &mut pivot_caches,
                            next_cache_id: &mut next_cache_id,
                        },
                    )?;
                }
                Event::Empty(start) => {
                    handle_spreadsheet_empty_full(
                        start,
                        FullSpreadsheetState {
                            in_spreadsheet: &mut in_spreadsheet,
                            sheet_id: &mut sheet_id,
                            store,
                            limits,
                            validations: &mut validations,
                            sheets: &mut sheets,
                            sheet_index: &mut sheet_index,
                            pivot_links: &mut pivot_links,
                            pivot_caches: &mut pivot_caches,
                            next_cache_id: &mut next_cache_id,
                        },
                    );
                }
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        },
    )?;

    apply_pivot_links(store, &sheet_index, pivot_links);
    Ok(OdfContentResult {
        content: sheets,
        pivot_caches,
        ..OdfContentResult::default()
    })
}

pub(crate) fn parse_content_spreadsheet_fast(
    xml: &[u8],
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<OdfContentResult, ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut in_spreadsheet = false;
    let mut sheet_id = 1u32;
    let mut sheets = Vec::new();
    let validations: HashMap<String, ValidationDef> = HashMap::new();

    scan_xml_events_until_end(
        &mut reader,
        &mut buf,
        "content.xml",
        |event| is_end_event(event, b"office:spreadsheet"),
        |reader, event| {
            match event {
                Event::Start(start) => {
                    handle_spreadsheet_start_fast(
                        reader,
                        start,
                        FastSpreadsheetState {
                            in_spreadsheet: &mut in_spreadsheet,
                            sheet_id: &mut sheet_id,
                            store,
                            sheets: &mut sheets,
                            validations: &validations,
                            limits,
                        },
                    )?;
                }
                Event::Empty(start) => {
                    handle_spreadsheet_empty_fast(
                        start,
                        in_spreadsheet,
                        &mut sheet_id,
                        store,
                        &mut sheets,
                    );
                }
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        },
    )?;

    Ok(OdfContentResult {
        content: sheets,
        ..OdfContentResult::default()
    })
}

fn build_empty_sheet(start: &BytesStart<'_>, sheet_id: u32) -> (Worksheet, NodeId) {
    let name = attr_value(start, b"table:name").unwrap_or_else(|| format!("Sheet{}", sheet_id));
    let sheet = Worksheet::new(name, sheet_id);
    let node_id = sheet.id;
    (sheet, node_id)
}

fn handle_spreadsheet_start_full(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    state: FullSpreadsheetState<'_>,
) -> Result<(), ParseError> {
    let FullSpreadsheetState {
        in_spreadsheet,
        sheet_id,
        store,
        limits,
        validations,
        sheets,
        sheet_index,
        pivot_links,
        pivot_caches,
        next_cache_id,
    } = state;
    match start.name().as_ref() {
        b"office:spreadsheet" => *in_spreadsheet = true,
        b"table:content-validation" if *in_spreadsheet => {
            insert_validation_definition(validations, start);
        }
        b"table:table" if *in_spreadsheet => {
            let worksheet = parse_ods_table(reader, start, *sheet_id, store, validations, limits)?;
            insert_worksheet(store, sheets, sheet_index, worksheet);
            *sheet_id += 1;
        }
        b"table:data-pilot-table" if *in_spreadsheet => {
            let parsed = parse_ods_pivot_table_full(reader, start, *next_cache_id)?;
            record_pivot_parse(store, pivot_links, pivot_caches, next_cache_id, parsed);
        }
        _ => {}
    }
    Ok(())
}

fn handle_spreadsheet_empty_full(empty: &BytesStart<'_>, state: FullSpreadsheetState<'_>) {
    let FullSpreadsheetState {
        in_spreadsheet,
        sheet_id,
        store,
        validations,
        sheets,
        sheet_index,
        pivot_links,
        pivot_caches,
        next_cache_id,
        ..
    } = state;
    match empty.name().as_ref() {
        b"table:table" if *in_spreadsheet => {
            let (sheet, _) = build_empty_sheet(empty, *sheet_id);
            insert_worksheet(store, sheets, sheet_index, sheet);
            *sheet_id += 1;
        }
        b"table:content-validation" if *in_spreadsheet => {
            insert_validation_definition(validations, empty);
        }
        b"table:data-pilot-table" if *in_spreadsheet => {
            let parsed = parse_ods_pivot_table_empty(empty, *next_cache_id);
            record_pivot_parse(store, pivot_links, pivot_caches, next_cache_id, parsed);
        }
        _ => {}
    }
}

fn handle_spreadsheet_start_fast(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    state: FastSpreadsheetState<'_>,
) -> Result<(), ParseError> {
    let FastSpreadsheetState {
        in_spreadsheet,
        sheet_id,
        store,
        sheets,
        validations,
        limits,
    } = state;
    match start.name().as_ref() {
        b"office:spreadsheet" => *in_spreadsheet = true,
        b"table:table" if *in_spreadsheet => {
            let worksheet =
                parse_ods_table_fast(reader, start, *sheet_id, store, validations, limits)?;
            insert_worksheet_no_index(store, sheets, worksheet);
            *sheet_id += 1;
        }
        _ => {}
    }
    Ok(())
}

fn handle_spreadsheet_empty_fast(
    empty: &BytesStart<'_>,
    in_spreadsheet: bool,
    sheet_id: &mut u32,
    store: &mut IrStore,
    sheets: &mut Vec<NodeId>,
) {
    if empty.name().as_ref() == b"table:table" && in_spreadsheet {
        let (sheet, _) = build_empty_sheet(empty, *sheet_id);
        insert_worksheet_no_index(store, sheets, sheet);
        *sheet_id += 1;
    }
}

fn insert_validation_definition(
    validations: &mut HashMap<String, ValidationDef>,
    element: &BytesStart<'_>,
) {
    if let Some((name, def)) = parse_validation_definition(element) {
        validations.insert(name, def);
    }
}

fn insert_worksheet(
    store: &mut IrStore,
    sheets: &mut Vec<NodeId>,
    sheet_index: &mut HashMap<String, NodeId>,
    worksheet: Worksheet,
) {
    let node_id = worksheet.id;
    sheet_index.insert(worksheet.name.clone(), node_id);
    store.insert(IRNode::Worksheet(worksheet));
    sheets.push(node_id);
}

fn insert_worksheet_no_index(store: &mut IrStore, sheets: &mut Vec<NodeId>, worksheet: Worksheet) {
    let node_id = worksheet.id;
    store.insert(IRNode::Worksheet(worksheet));
    sheets.push(node_id);
}

fn apply_pivot_links(
    store: &mut IrStore,
    sheet_index: &HashMap<String, NodeId>,
    pivot_links: Vec<(Option<String>, NodeId)>,
) {
    for (sheet_name, pivot_id) in pivot_links {
        if let Some(name) = sheet_name {
            if let Some(sheet_id) = sheet_index.get(&name).copied() {
                if let Some(IRNode::Worksheet(sheet)) = store.get_mut(sheet_id) {
                    sheet.pivot_tables.push(pivot_id);
                }
            }
        }
    }
}

pub(crate) fn parse_content_spreadsheet_parallel(
    xml: &[u8],
    store: &mut IrStore,
    limits: &Arc<OdfAtomicLimits>,
    config: &ParserConfig,
) -> Result<OdfContentResult, ParseError> {
    let validations = Arc::new(collect_validation_definitions(xml)?);
    let chunks = extract_spreadsheet_table_chunks(xml);
    if chunks.is_empty() {
        return Ok(OdfContentResult::default());
    }

    let max_threads = config
        .odf
        .parallel_max_threads
        .or_else(|| thread::available_parallelism().ok().map(|n| n.get()))
        .unwrap_or(1)
        .max(1);

    let total = chunks.len();
    let mut results: Vec<Option<Result<OdfSheetParseResult, ParseError>>> =
        (0..total).map(|_| None).collect();

    let mut batch_start = 0usize;
    while batch_start < total {
        let batch_end = (batch_start + max_threads).min(total);
        let (tx, rx) = std::sync::mpsc::channel();
        for (offset, chunk) in chunks[batch_start..batch_end].iter().cloned().enumerate() {
            let idx = batch_start + offset;
            let tx = tx.clone();
            let validations: Arc<HashMap<String, ValidationDef>> = Arc::clone(&validations);
            let limits = Arc::clone(limits);
            let sheet_id = (idx + 1) as u32;
            thread::spawn(move || {
                let result = parse_ods_table_from_chunk(
                    &chunk,
                    sheet_id,
                    validations.as_ref(),
                    limits.as_ref(),
                );
                let _ = tx.send((idx, result));
            });
        }
        for _ in batch_start..batch_end {
            if let Ok((i, res)) = rx.recv() {
                results[i] = Some(res);
            }
        }
        batch_start = batch_end;
    }

    let mut sheets = Vec::new();
    let mut sheet_index: HashMap<String, NodeId> = HashMap::new();
    for res in results {
        let Some(res) = res else {
            continue;
        };
        let parsed = res?;
        for node in parsed.nodes {
            store.insert(node);
        }
        let sheet_id = parsed.worksheet.id;
        sheet_index.insert(parsed.worksheet.name.clone(), sheet_id);
        store.insert(IRNode::Worksheet(parsed.worksheet));
        sheets.push(sheet_id);
    }

    let (pivot_links, pivot_caches) = parse_ods_pivots_from_xml(xml, store)?;
    apply_pivot_links(store, &sheet_index, pivot_links);

    Ok(OdfContentResult {
        content: sheets,
        pivot_caches,
        ..OdfContentResult::default()
    })
}

struct OdfSheetParseResult {
    worksheet: Worksheet,
    nodes: Vec<IRNode>,
}

fn parse_ods_table_from_chunk(
    chunk: &OdfTableChunk,
    sheet_id: u32,
    validations: &HashMap<String, ValidationDef>,
    limits: &OdfAtomicLimits,
) -> Result<OdfSheetParseResult, ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(chunk.bytes.as_slice()));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut parsed = None;
    scan_xml_events_with_reader(&mut reader, &mut buf, "content.xml", |reader, event| {
        if let Event::Start(e) = event {
            if e.name().as_ref() == b"table:table" {
                let mut local_store = IrStore::new();
                let worksheet =
                    parse_ods_table(reader, &e, sheet_id, &mut local_store, validations, limits)?;
                parsed = Some(OdfSheetParseResult {
                    worksheet,
                    nodes: local_store.into_nodes(),
                });
                return Ok(XmlScanControl::Break);
            }
        }
        Ok(XmlScanControl::Continue)
    })?;

    if let Some(parsed) = parsed {
        return Ok(parsed);
    }

    let name = table_name_from_chunk(&chunk.bytes, sheet_id);
    Ok(OdfSheetParseResult {
        worksheet: Worksheet::new(name, sheet_id),
        nodes: Vec::new(),
    })
}

pub(crate) fn skip_element(reader: &mut OdfReader<'_>, end: &[u8]) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    let mut depth: usize = 0;
    scan_xml_events(reader, &mut buf, "content.xml", |event| {
        match event {
            Event::Start(_) => {
                depth += 1;
            }
            Event::End(e) if e.name().as_ref() == end && depth == 0 => {
                return Ok(XmlScanControl::Break);
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;
    Ok(())
}
