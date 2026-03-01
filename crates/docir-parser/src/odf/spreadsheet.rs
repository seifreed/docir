//! ODF spreadsheet parsing helpers.

use super::presentation_helpers::{
    parse_frame_shape_empty, parse_frame_shape_start, FrameShapeState,
};
use super::spreadsheet_chunks::{
    extract_spreadsheet_table_chunks, table_name_from_chunk, OdfTableChunk,
};
use super::*;

pub(super) fn parse_content_spreadsheet(
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
    let mut pivot_links: Vec<(Option<String>, NodeId)> = Vec::new();
    let mut pivot_caches: Vec<NodeId> = Vec::new();
    let mut next_cache_id: u32 = 1;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => handle_spreadsheet_start_full(
                &mut reader,
                &e,
                &mut in_spreadsheet,
                &mut sheet_id,
                store,
                limits,
                &mut validations,
                &mut sheets,
                &mut sheet_index,
                &mut pivot_links,
                &mut pivot_caches,
                &mut next_cache_id,
            )?,
            Ok(Event::Empty(e)) => handle_spreadsheet_empty_full(
                &e,
                in_spreadsheet,
                &mut sheet_id,
                store,
                &mut validations,
                &mut sheets,
                &mut sheet_index,
                &mut pivot_links,
                &mut pivot_caches,
                &mut next_cache_id,
            ),
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"office:spreadsheet" {
                    in_spreadsheet = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(content_xml_error(e));
            }
            _ => {}
        }
        buf.clear();
    }

    let mut result = OdfContentResult::default();
    result.content = sheets;
    result.pivot_caches = pivot_caches;
    apply_pivot_links(store, &sheet_index, pivot_links);
    Ok(result)
}

pub(super) fn parse_draw_frame_spreadsheet(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let transform = parse_frame_transform(start);
    let mut frame = FrameShapeState::new();
    let mut buf = Vec::new();
    let mut name = attr_value(start, b"draw:name");

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => parse_frame_shape_start(reader, &e, store, &mut frame)?,
            Ok(Event::Empty(e)) => parse_frame_shape_empty(&e, store, &mut frame),
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"draw:frame" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error("content.xml", e)),
            _ => {}
        }
        buf.clear();
    }

    if frame.has_shape {
        let mut shape = Shape::new(frame.shape_type);
        shape.name = name.take();
        shape.media_target = frame.media_target;
        shape.chart_id = frame.chart_id;
        shape.transform = transform;
        let shape_id = shape.id;
        store.insert(IRNode::Shape(shape));
        Ok(Some(shape_id))
    } else {
        Ok(None)
    }
}

pub(super) fn parse_content_spreadsheet_fast(
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

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => handle_spreadsheet_start_fast(
                &mut reader,
                &e,
                &mut in_spreadsheet,
                &mut sheet_id,
                store,
                &validations,
                limits,
                &mut sheets,
            )?,
            Ok(Event::Empty(e)) => {
                handle_spreadsheet_empty_fast(&e, in_spreadsheet, &mut sheet_id, store, &mut sheets)
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"office:spreadsheet" {
                    in_spreadsheet = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(content_xml_error(e));
            }
            _ => {}
        }
        buf.clear();
    }

    let mut result = OdfContentResult::default();
    result.content = sheets;
    Ok(result)
}

fn build_empty_sheet(start: &BytesStart<'_>, sheet_id: u32) -> (Worksheet, NodeId) {
    let name = attr_value(start, b"table:name").unwrap_or_else(|| format!("Sheet{}", sheet_id));
    let sheet = Worksheet::new(name, sheet_id);
    let node_id = sheet.id;
    (sheet, node_id)
}

#[allow(clippy::too_many_arguments)]
fn handle_spreadsheet_start_full(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    in_spreadsheet: &mut bool,
    sheet_id: &mut u32,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    validations: &mut HashMap<String, ValidationDef>,
    sheets: &mut Vec<NodeId>,
    sheet_index: &mut HashMap<String, NodeId>,
    pivot_links: &mut Vec<(Option<String>, NodeId)>,
    pivot_caches: &mut Vec<NodeId>,
    next_cache_id: &mut u32,
) -> Result<(), ParseError> {
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

#[allow(clippy::too_many_arguments)]
fn handle_spreadsheet_empty_full(
    empty: &BytesStart<'_>,
    in_spreadsheet: bool,
    sheet_id: &mut u32,
    store: &mut IrStore,
    validations: &mut HashMap<String, ValidationDef>,
    sheets: &mut Vec<NodeId>,
    sheet_index: &mut HashMap<String, NodeId>,
    pivot_links: &mut Vec<(Option<String>, NodeId)>,
    pivot_caches: &mut Vec<NodeId>,
    next_cache_id: &mut u32,
) {
    match empty.name().as_ref() {
        b"table:table" if in_spreadsheet => {
            let (sheet, _) = build_empty_sheet(empty, *sheet_id);
            insert_worksheet(store, sheets, sheet_index, sheet);
            *sheet_id += 1;
        }
        b"table:content-validation" if in_spreadsheet => {
            insert_validation_definition(validations, empty);
        }
        b"table:data-pilot-table" if in_spreadsheet => {
            let parsed = parse_ods_pivot_table_empty(empty, *next_cache_id);
            record_pivot_parse(store, pivot_links, pivot_caches, next_cache_id, parsed);
        }
        _ => {}
    }
}

fn handle_spreadsheet_start_fast(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    in_spreadsheet: &mut bool,
    sheet_id: &mut u32,
    store: &mut IrStore,
    validations: &HashMap<String, ValidationDef>,
    limits: &dyn OdfLimitCounter,
    sheets: &mut Vec<NodeId>,
) -> Result<(), ParseError> {
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

fn content_xml_error(err: quick_xml::Error) -> ParseError {
    ParseError::Xml {
        file: "content.xml".to_string(),
        message: err.to_string(),
    }
}

fn parse_ods_pivot_table_full(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    cache_id: u32,
) -> Result<
    Option<(
        PivotTable,
        Option<String>,
        Option<PivotCache>,
        Option<PivotCacheRecords>,
    )>,
    ParseError,
> {
    let (mut pivot, sheet_name, cache) = build_ods_pivot(start, cache_id);
    let mut field_count: u32 = 0;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"table:data-pilot-field" {
                    field_count = field_count.saturating_add(1);
                    skip_element(reader, e.name().as_ref())?;
                }
            }
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"table:data-pilot-field" {
                    field_count = field_count.saturating_add(1);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:data-pilot-table" {
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

    let mut records: Option<PivotCacheRecords> = None;
    if let Some(cache) = cache.as_ref() {
        if field_count > 0 {
            let mut rec = PivotCacheRecords::new();
            rec.cache_id = Some(cache.cache_id);
            rec.field_count = Some(field_count);
            rec.span = Some(SourceSpan::new("content.xml"));
            records = Some(rec);
        }
        pivot.cache_id = Some(cache.cache_id);
    }

    Ok(Some((pivot, sheet_name, cache, records)))
}

fn parse_ods_pivot_table_empty(
    start: &BytesStart<'_>,
    cache_id: u32,
) -> Option<(
    PivotTable,
    Option<String>,
    Option<PivotCache>,
    Option<PivotCacheRecords>,
)> {
    let (mut pivot, sheet_name, cache) = build_ods_pivot(start, cache_id);
    if let Some(cache) = cache.as_ref() {
        pivot.cache_id = Some(cache.cache_id);
    }
    Some((pivot, sheet_name, cache, None))
}

fn record_pivot_parse(
    store: &mut IrStore,
    pivot_links: &mut Vec<(Option<String>, NodeId)>,
    pivot_caches: &mut Vec<NodeId>,
    next_cache_id: &mut u32,
    parsed: Option<(
        PivotTable,
        Option<String>,
        Option<PivotCache>,
        Option<PivotCacheRecords>,
    )>,
) {
    let Some((pivot, sheet_name, cache, records)) = parsed else {
        return;
    };

    let id = pivot.id;
    store.insert(IRNode::PivotTable(pivot));
    pivot_links.push((sheet_name, id));

    if let Some(mut cache) = cache {
        if let Some(records) = records {
            let records_id = records.id;
            store.insert(IRNode::PivotCacheRecords(records));
            cache.records = Some(records_id);
        }
        let cache_node_id = cache.id;
        store.insert(IRNode::PivotCache(cache));
        pivot_caches.push(cache_node_id);
        *next_cache_id = next_cache_id.saturating_add(1);
    }
}

fn build_ods_pivot(
    start: &BytesStart<'_>,
    cache_id: u32,
) -> (PivotTable, Option<String>, Option<PivotCache>) {
    let name = attr_value(start, b"table:name");
    let target = attr_value(start, b"table:target-range-address");
    let source = attr_value(start, b"table:source-range-address");
    let ref_range = target.clone().or(source.clone());
    let sheet_name = ref_range
        .as_deref()
        .and_then(extract_sheet_name)
        .map(|s| s.to_string());
    let mut pivot = PivotTable {
        id: NodeId::new(),
        name,
        cache_id: None,
        ref_range,
        span: Some(SourceSpan::new("content.xml")),
    };
    if pivot.name.is_none() {
        pivot.name = attr_value(start, b"table:display-name");
    }
    let cache = source.map(|source_range| {
        let mut cache = PivotCache::new(cache_id);
        cache.cache_source = Some(source_range);
        cache.span = Some(SourceSpan::new("content.xml"));
        cache
    });
    (pivot, sheet_name, cache)
}

fn extract_sheet_name(range: &str) -> Option<&str> {
    let trimmed = range.trim();
    let (name_part, _) = trimmed.split_once('.')?;
    Some(name_part.trim_matches('\''))
}

pub(super) fn parse_content_spreadsheet_parallel(
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
            let validations = Arc::clone(&validations);
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
    for (sheet_name, pivot_id) in pivot_links {
        if let Some(name) = sheet_name {
            if let Some(sheet_id) = sheet_index.get(&name).copied() {
                if let Some(IRNode::Worksheet(sheet)) = store.get_mut(sheet_id) {
                    sheet.pivot_tables.push(pivot_id);
                }
            }
        }
    }

    let mut result = OdfContentResult::default();
    result.content = sheets;
    result.pivot_caches = pivot_caches;
    Ok(result)
}

fn parse_ods_pivots_from_xml(
    xml: &[u8],
    store: &mut IrStore,
) -> Result<(Vec<(Option<String>, NodeId)>, Vec<NodeId>), ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut in_spreadsheet = false;
    let mut pivot_links: Vec<(Option<String>, NodeId)> = Vec::new();
    let mut pivot_caches: Vec<NodeId> = Vec::new();
    let mut next_cache_id: u32 = 1;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"office:spreadsheet" => in_spreadsheet = true,
                b"table:data-pilot-table" if in_spreadsheet => {
                    let cache_id = next_cache_id;
                    let parsed = parse_ods_pivot_table_full(&mut reader, &e, cache_id)?;
                    record_pivot_parse(
                        store,
                        &mut pivot_links,
                        &mut pivot_caches,
                        &mut next_cache_id,
                        parsed,
                    );
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:data-pilot-table" if in_spreadsheet => {
                    let cache_id = next_cache_id;
                    let parsed = parse_ods_pivot_table_empty(&e, cache_id);
                    record_pivot_parse(
                        store,
                        &mut pivot_links,
                        &mut pivot_caches,
                        &mut next_cache_id,
                        parsed,
                    );
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"office:spreadsheet" {
                    in_spreadsheet = false;
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

    Ok((pivot_links, pivot_caches))
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
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"table:table" => {
                let mut local_store = IrStore::new();
                let worksheet = parse_ods_table(
                    &mut reader,
                    &e,
                    sheet_id,
                    &mut local_store,
                    validations,
                    limits,
                )?;
                return Ok(OdfSheetParseResult {
                    worksheet,
                    nodes: local_store.into_nodes(),
                });
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

    let name = table_name_from_chunk(&chunk.bytes, sheet_id);
    Ok(OdfSheetParseResult {
        worksheet: Worksheet::new(name, sheet_id),
        nodes: Vec::new(),
    })
}

fn collect_validation_definitions(
    xml: &[u8],
) -> Result<HashMap<String, ValidationDef>, ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut in_spreadsheet = false;
    let mut validations: HashMap<String, ValidationDef> = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"office:spreadsheet" => in_spreadsheet = true,
                b"table:content-validation" if in_spreadsheet => {
                    if let Some((name, def)) = parse_validation_definition(&e) {
                        validations.insert(name, def);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:content-validation" if in_spreadsheet => {
                    if let Some((name, def)) = parse_validation_definition(&e) {
                        validations.insert(name, def);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"office:spreadsheet" {
                    in_spreadsheet = false;
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

    Ok(validations)
}

pub(super) fn skip_element(reader: &mut OdfReader<'_>, end: &[u8]) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    let mut depth: usize = 0;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(_)) => {
                depth += 1;
            }
            Ok(Event::End(e)) => {
                if depth == 0 && e.name().as_ref() == end {
                    break;
                }
                depth = depth.saturating_sub(1);
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
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn default_limits() -> OdfLimits {
        OdfLimits::new(&ParserConfig::default(), false)
    }

    fn fast_limits() -> OdfLimits {
        OdfLimits::new(&ParserConfig::default(), true)
    }

    #[test]
    fn parse_content_spreadsheet_links_pivots_and_validations() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <table:table table:name="OutsideSpreadsheet"/>
    <office:spreadsheet>
      <table:content-validations>
        <table:content-validation table:name="val1" table:condition="cell-content-is-between(1,10)" />
      </table:content-validations>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell table:cell-value-type="float" table:cell-value="1" table:content-validation-name="val1" />
        </table:table-row>
      </table:table>
      <table:table table:name="Sheet2"/>
      <table:data-pilot-table table:name="Pivot1"
        table:source-range-address="'Sheet1'.A1:'Sheet1'.A1"
        table:target-range-address="'Sheet1'.D1:'Sheet1'.D1">
        <table:data-pilot-field table:source-field-name="Field1"/>
      </table:data-pilot-table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;

        let mut store = IrStore::new();
        let result = parse_content_spreadsheet(xml, &mut store, &default_limits()).unwrap();

        assert_eq!(result.content.len(), 2);
        assert_eq!(result.pivot_caches.len(), 1);

        let sheet1 = result.content[0];
        let sheet2 = result.content[1];
        let Some(IRNode::Worksheet(sheet1_node)) = store.get(sheet1) else {
            panic!("expected first content node to be worksheet");
        };
        let Some(IRNode::Worksheet(sheet2_node)) = store.get(sheet2) else {
            panic!("expected second content node to be worksheet");
        };

        assert_eq!(sheet1_node.name, "Sheet1");
        assert_eq!(sheet2_node.name, "Sheet2");
        assert_eq!(sheet1_node.pivot_tables.len(), 1);
        assert!(sheet2_node.pivot_tables.is_empty());

        let validation_count = store
            .values()
            .filter(|node| matches!(node, IRNode::DataValidation(_)))
            .count();
        assert_eq!(validation_count, 1);
    }

    #[test]
    fn parse_content_spreadsheet_reports_xml_errors() {
        let malformed = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:spreadsheet><table:table xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"></office:spreadsheet></office:body></office:document-content>"#;

        let mut store = IrStore::new();
        let err = parse_content_spreadsheet(malformed, &mut store, &default_limits()).unwrap_err();
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
            other => panic!("expected xml error, got {:?}", other),
        }
    }

    #[test]
    fn parse_content_spreadsheet_fast_emits_default_sheet_without_pivot_cache() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table/>
      <table:data-pilot-table table:name="PivotIgnored"
        table:source-range-address="'Sheet1'.A1:'Sheet1'.A1"
        table:target-range-address="'Sheet1'.D1:'Sheet1'.D1"/>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;

        let mut store = IrStore::new();
        let result = parse_content_spreadsheet(xml, &mut store, &fast_limits()).unwrap();

        assert_eq!(result.content.len(), 1);
        assert!(result.pivot_caches.is_empty());

        let Some(IRNode::Worksheet(sheet)) = store.get(result.content[0]) else {
            panic!("expected worksheet node");
        };
        assert_eq!(sheet.name, "Sheet1");
        assert!(sheet.pivot_tables.is_empty());
    }

    #[test]
    fn parse_content_spreadsheet_fast_reports_xml_errors() {
        let malformed = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:spreadsheet><table:table xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"></office:spreadsheet></office:body></office:document-content>"#;

        let mut store = IrStore::new();
        let err = parse_content_spreadsheet(malformed, &mut store, &fast_limits()).unwrap_err();
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
            other => panic!("expected xml error, got {:?}", other),
        }
    }

    #[test]
    fn parse_content_spreadsheet_assigns_default_names_for_empty_tables() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table/>
      <table:table table:name="Named"/>
      <table:table/>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;

        let mut store = IrStore::new();
        let result = parse_content_spreadsheet(xml, &mut store, &default_limits()).unwrap();
        assert_eq!(result.content.len(), 3);

        let Some(IRNode::Worksheet(first)) = store.get(result.content[0]) else {
            panic!("expected first worksheet");
        };
        let Some(IRNode::Worksheet(second)) = store.get(result.content[1]) else {
            panic!("expected second worksheet");
        };
        let Some(IRNode::Worksheet(third)) = store.get(result.content[2]) else {
            panic!("expected third worksheet");
        };
        assert_eq!(first.name, "Sheet1");
        assert_eq!(second.name, "Named");
        assert_eq!(third.name, "Sheet3");
    }

    #[test]
    fn parse_content_spreadsheet_keeps_unlinked_pivot_cache_without_records() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1"/>
      <table:data-pilot-table table:display-name="PivotDisplay"
        table:source-range-address="'Missing'.A1:'Missing'.A1"
        table:target-range-address="'Missing'.D1:'Missing'.D1"/>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;

        let mut store = IrStore::new();
        let result = parse_content_spreadsheet(xml, &mut store, &default_limits()).unwrap();
        assert_eq!(result.content.len(), 1);
        assert_eq!(result.pivot_caches.len(), 1);

        let Some(IRNode::Worksheet(sheet)) = store.get(result.content[0]) else {
            panic!("expected worksheet node");
        };
        assert_eq!(sheet.name, "Sheet1");
        assert!(sheet.pivot_tables.is_empty());

        let pivots: Vec<_> = store
            .values()
            .filter_map(|node| match node {
                IRNode::PivotTable(pivot) => Some(pivot),
                _ => None,
            })
            .collect();
        assert_eq!(pivots.len(), 1);
        assert_eq!(pivots[0].name.as_deref(), Some("PivotDisplay"));
        assert_eq!(
            pivots[0].ref_range.as_deref(),
            Some("'Missing'.D1:'Missing'.D1")
        );

        let Some(IRNode::PivotCache(cache)) = store.get(result.pivot_caches[0]) else {
            panic!("expected pivot cache node");
        };
        assert_eq!(
            cache.cache_source.as_deref(),
            Some("'Missing'.A1:'Missing'.A1")
        );
        assert!(cache.records.is_none());
    }

    #[test]
    fn parse_draw_frame_spreadsheet_emits_shape_for_plugin_media() {
        let xml = br#"<draw:frame xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
                xmlns:xlink="http://www.w3.org/1999/xlink"
                draw:name="MediaFrame">
            <draw:plugin xlink:href="media/movie.mp4"/>
        </draw:frame>"#;
        let mut reader = Reader::from_reader(std::io::Cursor::new(xml.as_slice()));
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let frame_start = loop {
            match reader.read_event_into(&mut buf).unwrap() {
                Event::Start(e) if e.name().as_ref() == b"draw:frame" => break e.into_owned(),
                Event::Eof => panic!("missing draw:frame"),
                _ => {}
            }
            buf.clear();
        };
        let mut store = IrStore::new();

        let shape_id = parse_draw_frame_spreadsheet(&mut reader, &frame_start, &mut store).unwrap();
        let Some(shape_id) = shape_id else {
            panic!("expected shape");
        };
        let Some(IRNode::Shape(shape)) = store.get(shape_id) else {
            panic!("expected shape node");
        };
        assert_eq!(shape.name.as_deref(), Some("MediaFrame"));
        assert_eq!(shape.media_target.as_deref(), Some("media/movie.mp4"));
        assert_eq!(shape.shape_type, ShapeType::Video);
    }

    #[test]
    fn parse_draw_frame_spreadsheet_returns_none_without_shape_payload() {
        let xml = br#"<draw:frame xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
                draw:name="EmptyFrame">
            <draw:unknown/>
        </draw:frame>"#;
        let mut reader = Reader::from_reader(std::io::Cursor::new(xml.as_slice()));
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let frame_start = loop {
            match reader.read_event_into(&mut buf).unwrap() {
                Event::Start(e) if e.name().as_ref() == b"draw:frame" => break e.into_owned(),
                Event::Eof => panic!("missing draw:frame"),
                _ => {}
            }
            buf.clear();
        };
        let mut store = IrStore::new();

        let shape_id = parse_draw_frame_spreadsheet(&mut reader, &frame_start, &mut store).unwrap();
        assert!(shape_id.is_none());
        assert_eq!(store.values().count(), 0);
    }

    #[test]
    fn parse_draw_frame_spreadsheet_returns_none_on_truncated_frame() {
        let xml = br#"<draw:frame xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><draw:plugin>"#;
        let mut reader = Reader::from_reader(std::io::Cursor::new(xml.as_slice()));
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let frame_start = loop {
            match reader.read_event_into(&mut buf).unwrap() {
                Event::Start(e) if e.name().as_ref() == b"draw:frame" => break e.into_owned(),
                Event::Eof => panic!("missing draw:frame"),
                _ => {}
            }
            buf.clear();
        };
        let mut store = IrStore::new();
        let shape_id = parse_draw_frame_spreadsheet(&mut reader, &frame_start, &mut store).unwrap();
        assert!(shape_id.is_none());
        assert_eq!(store.values().count(), 0);
    }

    #[test]
    fn parse_content_spreadsheet_parallel_links_pivots_and_caches() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:content-validations>
        <table:content-validation table:name="v1" table:condition="cell-content-is-between(1,2)"/>
      </table:content-validations>
      <table:table table:name="SheetA">
        <table:table-row>
          <table:table-cell table:content-validation-name="v1"/>
        </table:table-row>
      </table:table>
      <table:table table:name="SheetB"/>
      <table:data-pilot-table table:name="PivotA"
        table:source-range-address="'SheetA'.A1:'SheetA'.A1"
        table:target-range-address="'SheetA'.C1:'SheetA'.C1">
        <table:data-pilot-field table:source-field-name="Amount"/>
      </table:data-pilot-table>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;

        let mut config = ParserConfig::default();
        config.odf.parallel_max_threads = Some(2);
        let limits = std::sync::Arc::new(OdfAtomicLimits::new(&config, false));
        let mut store = IrStore::new();

        let result = parse_content_spreadsheet_parallel(xml, &mut store, &limits, &config).unwrap();
        assert_eq!(result.content.len(), 2);
        assert_eq!(result.pivot_caches.len(), 1);

        let Some(IRNode::Worksheet(sheet_a)) = store.get(result.content[0]) else {
            panic!("expected worksheet");
        };
        assert_eq!(sheet_a.name, "SheetA");
        assert_eq!(sheet_a.pivot_tables.len(), 1);

        let Some(IRNode::PivotTable(pivot)) = store.get(sheet_a.pivot_tables[0]) else {
            panic!("expected pivot table");
        };
        assert_eq!(pivot.name.as_deref(), Some("PivotA"));
        assert_eq!(pivot.ref_range.as_deref(), Some("'SheetA'.C1:'SheetA'.C1"));
        assert!(pivot.cache_id.is_some());

        let Some(IRNode::PivotCache(cache)) = store.get(result.pivot_caches[0]) else {
            panic!("expected pivot cache");
        };
        assert_eq!(
            cache.cache_source.as_deref(),
            Some("'SheetA'.A1:'SheetA'.A1")
        );
        assert!(cache.records.is_some());
    }

    #[test]
    fn parse_content_spreadsheet_parallel_handles_missing_sheet_chunks() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:content-validations>
        <table:content-validation table:name="v1" table:condition="cell-content-is-whole-number()"/>
      </table:content-validations>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;

        let config = ParserConfig::default();
        let limits = std::sync::Arc::new(OdfAtomicLimits::new(&config, false));
        let mut store = IrStore::new();

        let result = parse_content_spreadsheet_parallel(xml, &mut store, &limits, &config).unwrap();
        assert!(result.content.is_empty());
        assert!(result.pivot_caches.is_empty());
        assert_eq!(store.values().count(), 0);
    }

    #[test]
    fn parse_content_spreadsheet_parallel_handles_empty_pivot_without_cache() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="SheetA"><table:table-row/></table:table>
      <table:data-pilot-table table:display-name="PivotNoCache"
        table:target-range-address="'SheetA'.E1:'SheetA'.E1"/>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;

        let config = ParserConfig::default();
        let limits = std::sync::Arc::new(OdfAtomicLimits::new(&config, false));
        let mut store = IrStore::new();

        let result = parse_content_spreadsheet_parallel(xml, &mut store, &limits, &config).unwrap();
        assert_eq!(result.content.len(), 1);
        assert!(result.pivot_caches.is_empty());

        let Some(IRNode::Worksheet(sheet)) = store.get(result.content[0]) else {
            panic!("expected worksheet");
        };
        assert_eq!(sheet.name, "SheetA");
        assert_eq!(sheet.pivot_tables.len(), 1);

        let Some(IRNode::PivotTable(pivot)) = store.get(sheet.pivot_tables[0]) else {
            panic!("expected pivot table");
        };
        assert_eq!(pivot.name.as_deref(), Some("PivotNoCache"));
        assert_eq!(pivot.ref_range.as_deref(), Some("'SheetA'.E1:'SheetA'.E1"));
        assert!(pivot.cache_id.is_none());
    }

    #[test]
    fn parse_content_spreadsheet_parallel_reports_xml_error_for_malformed_pivot() {
        let malformed = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="SheetA"/>
    <table:data-pilot-table table:name="Broken"><table:data-pilot-field>
  </office:spreadsheet></office:body></office:document-content>"#;

        let config = ParserConfig::default();
        let limits = std::sync::Arc::new(OdfAtomicLimits::new(&config, false));
        let mut store = IrStore::new();
        let err = parse_content_spreadsheet_parallel(malformed, &mut store, &limits, &config)
            .unwrap_err();
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
            other => panic!("expected xml error, got {:?}", other),
        }
    }
}
