//! ODF spreadsheet parsing helpers.

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
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"office:spreadsheet" => in_spreadsheet = true,
                b"table:content-validation" if in_spreadsheet => {
                    if let Some((name, def)) = parse_validation_definition(&e) {
                        validations.insert(name, def);
                    }
                }
                b"table:table" if in_spreadsheet => {
                    let worksheet =
                        parse_ods_table(&mut reader, &e, sheet_id, store, &validations, limits)?;
                    let node_id = worksheet.id;
                    sheet_index.insert(worksheet.name.clone(), node_id);
                    store.insert(IRNode::Worksheet(worksheet));
                    sheets.push(node_id);
                    sheet_id += 1;
                }
                b"table:data-pilot-table" if in_spreadsheet => {
                    if let Some((pivot, sheet_name, cache, records)) =
                        parse_ods_pivot_table_full(&mut reader, &e, next_cache_id)?
                    {
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
                            next_cache_id = next_cache_id.saturating_add(1);
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table" if in_spreadsheet => {
                    let name = attr_value(&e, b"table:name")
                        .unwrap_or_else(|| format!("Sheet{}", sheet_id));
                    let sheet = Worksheet::new(name, sheet_id);
                    let node_id = sheet.id;
                    sheet_index.insert(sheet.name.clone(), node_id);
                    store.insert(IRNode::Worksheet(sheet));
                    sheets.push(node_id);
                    sheet_id += 1;
                }
                b"table:content-validation" if in_spreadsheet => {
                    if let Some((name, def)) = parse_validation_definition(&e) {
                        validations.insert(name, def);
                    }
                }
                b"table:data-pilot-table" if in_spreadsheet => {
                    if let Some((pivot, sheet_name, cache, records)) =
                        parse_ods_pivot_table_empty(&e, next_cache_id)
                    {
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
                            next_cache_id = next_cache_id.saturating_add(1);
                        }
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

    let mut result = OdfContentResult::default();
    result.content = sheets;
    result.pivot_caches = pivot_caches;
    for (sheet_name, pivot_id) in pivot_links {
        if let Some(name) = sheet_name {
            if let Some(sheet_id) = sheet_index.get(&name).copied() {
                if let Some(IRNode::Worksheet(sheet)) = store.get_mut(sheet_id) {
                    sheet.pivot_tables.push(pivot_id);
                }
            }
        }
    }
    Ok(result)
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
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"office:spreadsheet" => in_spreadsheet = true,
                b"table:table" if in_spreadsheet => {
                    let worksheet = parse_ods_table_fast(
                        &mut reader,
                        &e,
                        sheet_id,
                        store,
                        &validations,
                        limits,
                    )?;
                    let node_id = worksheet.id;
                    store.insert(IRNode::Worksheet(worksheet));
                    sheets.push(node_id);
                    sheet_id += 1;
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table" if in_spreadsheet => {
                    let name = attr_value(&e, b"table:name")
                        .unwrap_or_else(|| format!("Sheet{}", sheet_id));
                    let sheet = Worksheet::new(name, sheet_id);
                    let node_id = sheet.id;
                    store.insert(IRNode::Worksheet(sheet));
                    sheets.push(node_id);
                    sheet_id += 1;
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

    let mut result = OdfContentResult::default();
    result.content = sheets;
    Ok(result)
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
    let (mut pivot, sheet_name, mut cache) = build_ods_pivot(start, cache_id);
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
    let (mut pivot, sheet_name, mut cache) = build_ods_pivot(start, cache_id);
    if let Some(cache) = cache.as_ref() {
        pivot.cache_id = Some(cache.cache_id);
    }
    Some((pivot, sheet_name, cache, None))
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
        .odf_parallel_max_threads
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
                    if let Some((pivot, sheet_name, cache, records)) =
                        parse_ods_pivot_table_full(&mut reader, &e, next_cache_id)?
                    {
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
                            next_cache_id = next_cache_id.saturating_add(1);
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:data-pilot-table" if in_spreadsheet => {
                    if let Some((pivot, sheet_name, cache, records)) =
                        parse_ods_pivot_table_empty(&e, next_cache_id)
                    {
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
                            next_cache_id = next_cache_id.saturating_add(1);
                        }
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

    Ok((pivot_links, pivot_caches))
}

#[derive(Clone)]
pub(super) struct OdfTableChunk {
    pub(super) start: usize,
    pub(super) end: usize,
    pub(super) bytes: Vec<u8>,
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

pub(super) fn extract_spreadsheet_table_chunks(xml: &[u8]) -> Vec<OdfTableChunk> {
    let Some((start, end)) = find_spreadsheet_range(xml) else {
        return Vec::new();
    };
    let mut chunks = Vec::new();
    let mut pos = start;
    while let Some(idx) = find_subslice(xml, b"<table:table", pos, end) {
        if let Some(next) = xml.get(idx + b"<table:table".len()) {
            if *next == b'-' {
                pos = idx + 1;
                continue;
            }
        }
        let Some(tag_end) = find_tag_end(xml, idx + 1, end) else {
            break;
        };
        let self_closing = is_self_closing_tag(xml, idx, tag_end);
        let chunk_end = if self_closing {
            tag_end
        } else {
            let Some(close_start) = find_subslice(xml, b"</table:table>", tag_end + 1, end) else {
                break;
            };
            close_start + b"</table:table>".len() - 1
        };
        if chunk_end >= idx {
            let bytes = xml[idx..=chunk_end].to_vec();
            chunks.push(OdfTableChunk {
                start: idx,
                end: chunk_end,
                bytes,
            });
        }
        pos = chunk_end.saturating_add(1);
    }
    chunks
}

fn find_spreadsheet_range(xml: &[u8]) -> Option<(usize, usize)> {
    let start = find_subslice(xml, b"<office:spreadsheet", 0, xml.len())?;
    let tag_end = find_tag_end(xml, start + 1, xml.len())?;
    let end_tag = find_subslice(xml, b"</office:spreadsheet>", tag_end + 1, xml.len())?;
    let end = end_tag + b"</office:spreadsheet>".len();
    Some((tag_end + 1, end))
}

fn find_subslice(haystack: &[u8], needle: &[u8], start: usize, end: usize) -> Option<usize> {
    let mut i = start;
    let limit = end.saturating_sub(needle.len());
    while i <= limit {
        if &haystack[i..i + needle.len()] == needle {
            return Some(i);
        }
        i += 1;
    }
    None
}

fn find_tag_end(xml: &[u8], start: usize, end: usize) -> Option<usize> {
    let mut i = start;
    let mut quote: Option<u8> = None;
    while i < end {
        let b = xml[i];
        if let Some(q) = quote {
            if b == q {
                quote = None;
            }
        } else if b == b'"' || b == b'\'' {
            quote = Some(b);
        } else if b == b'>' {
            return Some(i);
        }
        i += 1;
    }
    None
}

fn is_self_closing_tag(xml: &[u8], start: usize, end: usize) -> bool {
    let mut i = end.saturating_sub(1);
    while i > start {
        let b = xml[i];
        if b == b'/' {
            return true;
        }
        if !b.is_ascii_whitespace() {
            break;
        }
        i = i.saturating_sub(1);
    }
    false
}

pub(super) fn table_name_from_chunk(chunk: &[u8], sheet_id: u32) -> String {
    let mut reader = Reader::from_reader(std::io::Cursor::new(chunk));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"table:table" => {
                return attr_value(&e, b"table:name")
                    .unwrap_or_else(|| format!("Sheet{}", sheet_id));
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
    format!("Sheet{}", sheet_id)
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
