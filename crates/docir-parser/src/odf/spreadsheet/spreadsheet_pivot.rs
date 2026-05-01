use super::super::super::helpers::ValidationDef;
use super::super::super::{attr_value, scan_xml_events_with_reader, OdfReader};
use crate::error::ParseError;
use crate::xml_utils::XmlScanControl;
use docir_core::ir::{IRNode, PivotCache, PivotCacheRecords, PivotTable};
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::HashMap;

type PivotParseResult = Option<(
    PivotTable,
    Option<String>,
    Option<PivotCache>,
    Option<PivotCacheRecords>,
)>;
pub(super) type PivotLinks = Vec<(Option<String>, NodeId)>;
pub(super) type PivotParseOutput = (PivotLinks, Vec<NodeId>);

pub(super) fn parse_ods_pivot_table_full(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    cache_id: u32,
) -> Result<PivotParseResult, ParseError> {
    let (mut pivot, sheet_name, cache) = build_ods_pivot(start, cache_id);
    let mut field_count: u32 = 0;
    let mut buf = Vec::new();
    scan_xml_events_with_reader(reader, &mut buf, "content.xml", |reader, event| {
        match event {
            Event::Start(e) => {
                if e.name().as_ref() == b"table:data-pilot-field" {
                    field_count = field_count.saturating_add(1);
                    super::skip_element(reader, e.name().as_ref())?;
                }
            }
            Event::Empty(e) => {
                if e.name().as_ref() == b"table:data-pilot-field" {
                    field_count = field_count.saturating_add(1);
                }
            }
            Event::End(e) if e.name().as_ref() == b"table:data-pilot-table" => {
                return Ok(XmlScanControl::Break);
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

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

pub(super) fn parse_ods_pivot_table_empty(
    start: &BytesStart<'_>,
    cache_id: u32,
) -> PivotParseResult {
    let (mut pivot, sheet_name, cache) = build_ods_pivot(start, cache_id);
    if let Some(cache) = cache.as_ref() {
        pivot.cache_id = Some(cache.cache_id);
    }
    Some((pivot, sheet_name, cache, None))
}

pub(super) fn record_pivot_parse(
    store: &mut IrStore,
    pivot_links: &mut PivotLinks,
    pivot_caches: &mut Vec<NodeId>,
    next_cache_id: &mut u32,
    parsed: PivotParseResult,
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

pub(super) fn parse_ods_pivots_from_xml(
    xml: &[u8],
    store: &mut IrStore,
) -> Result<PivotParseOutput, ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut in_spreadsheet = false;
    let mut pivot_links: PivotLinks = Vec::new();
    let mut pivot_caches: Vec<NodeId> = Vec::new();
    let mut next_cache_id: u32 = 1;

    scan_xml_events_with_reader(&mut reader, &mut buf, "content.xml", |reader, event| {
        match event {
            Event::Start(e) => match e.name().as_ref() {
                b"office:spreadsheet" => in_spreadsheet = true,
                b"table:data-pilot-table" if in_spreadsheet => {
                    let cache_id = next_cache_id;
                    let parsed = parse_ods_pivot_table_full(reader, &e, cache_id)?;
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
            Event::Empty(e) => {
                if e.name().as_ref() == b"table:data-pilot-table" && in_spreadsheet {
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
            }
            Event::End(e) if e.name().as_ref() == b"office:spreadsheet" => {
                in_spreadsheet = false;
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok((pivot_links, pivot_caches))
}

pub(super) fn collect_validation_definitions(
    xml: &[u8],
) -> Result<HashMap<String, ValidationDef>, ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut in_spreadsheet = false;
    let mut validations: HashMap<String, ValidationDef> = HashMap::new();

    super::scan_xml_events(&mut reader, &mut buf, "content.xml", |event| {
        match event {
            Event::Start(e) => match e.name().as_ref() {
                b"office:spreadsheet" => in_spreadsheet = true,
                b"table:content-validation" if in_spreadsheet => {
                    if let Some((name, def)) = super::parse_validation_definition(&e) {
                        validations.insert(name, def);
                    }
                }
                _ => {}
            },
            Event::Empty(e) => {
                if e.name().as_ref() == b"table:content-validation" && in_spreadsheet {
                    if let Some((name, def)) = super::parse_validation_definition(&e) {
                        validations.insert(name, def);
                    }
                }
            }
            Event::End(e) if e.name().as_ref() == b"office:spreadsheet" => {
                in_spreadsheet = false;
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok(validations)
}
