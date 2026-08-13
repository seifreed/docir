use super::super::super::helpers::ValidationDef;
use super::super::super::{OdfReader, scan_xml_events_with_reader};
use crate::error::ParseError;
use crate::xml_utils::{
    XmlScanControl, local_name, track_xml_root_event, try_attr_value_by_suffix,
};
use docir_core::ir::{IRNode, PivotCache, PivotCacheRecords, PivotTable};
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
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
    let (mut pivot, sheet_name, cache) = build_ods_pivot(start, cache_id)?;
    let mut field_count: u32 = 0;
    let mut buf = Vec::new();
    let mut reached_pivot_end = false;
    scan_xml_events_with_reader(reader, &mut buf, "content.xml", |reader, event| {
        match event {
            Event::Start(e) if local_name(e.name().as_ref()) == b"data-pilot-field" => {
                field_count = field_count.saturating_add(1);
                super::skip_element(reader, e.name().as_ref())?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"data-pilot-field" => {
                field_count = field_count.saturating_add(1);
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"data-pilot-table" => {
                reached_pivot_end = true;
                return Ok(XmlScanControl::Break);
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;
    if !reached_pivot_end {
        return Err(crate::xml_utils::xml_error(
            "content.xml",
            "unexpected end of XML before closing data-pilot-table",
        ));
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

pub(super) fn parse_ods_pivot_table_empty(
    start: &BytesStart<'_>,
    cache_id: u32,
) -> Result<PivotParseResult, ParseError> {
    let (mut pivot, sheet_name, cache) = build_ods_pivot(start, cache_id)?;
    if let Some(cache) = cache.as_ref() {
        pivot.cache_id = Some(cache.cache_id);
    }
    Ok(Some((pivot, sheet_name, cache, None)))
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
) -> Result<(PivotTable, Option<String>, Option<PivotCache>), ParseError> {
    let name = try_attr_value_by_suffix(start, &[b":name"], "content.xml")?;
    let target = try_attr_value_by_suffix(start, &[b":target-range-address"], "content.xml")?;
    let source = try_attr_value_by_suffix(start, &[b":source-range-address"], "content.xml")?;
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
        pivot.name = try_attr_value_by_suffix(start, &[b":display-name"], "content.xml")?;
    }
    let cache = source.map(|source_range| {
        let mut cache = PivotCache::new(cache_id);
        cache.cache_source = Some(source_range);
        cache.span = Some(SourceSpan::new("content.xml"));
        cache
    });
    Ok((pivot, sheet_name, cache))
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
    let mut root_name = None;
    let mut root_closed = false;

    scan_xml_events_with_reader(&mut reader, &mut buf, "content.xml", |reader, event| {
        track_xml_root_event(&event, &mut root_name, &mut root_closed, "content.xml")?;
        match event {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"spreadsheet" => in_spreadsheet = true,
                b"data-pilot-table" if in_spreadsheet => {
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
            Event::Empty(e)
                if local_name(e.name().as_ref()) == b"data-pilot-table" && in_spreadsheet =>
            {
                let cache_id = next_cache_id;
                let parsed = parse_ods_pivot_table_empty(&e, cache_id)?;
                record_pivot_parse(
                    store,
                    &mut pivot_links,
                    &mut pivot_caches,
                    &mut next_cache_id,
                    parsed,
                );
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"spreadsheet" => {
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
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"spreadsheet" => in_spreadsheet = true,
                b"content-validation" if in_spreadsheet => {
                    if let Some((name, def)) = super::parse_validation_definition(&e)? {
                        validations.insert(name, def);
                    }
                }
                _ => {}
            },
            Event::Empty(e) => {
                if local_name(e.name().as_ref()) == b"content-validation"
                    && in_spreadsheet
                    && let Some((name, def)) = super::parse_validation_definition(&e)?
                {
                    validations.insert(name, def);
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"spreadsheet" => {
                in_spreadsheet = false;
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok(validations)
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::events::Event;

    #[test]
    fn parse_ods_pivot_table_full_rejects_missing_end() {
        let xml: &[u8] = br#"<table:data-pilot-table
            xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            table:name="TruncatedPivot">"#;
        let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
        let mut buf = Vec::new();
        let start = loop {
            match reader.read_event_into(&mut buf).expect("pivot start") {
                Event::Start(event) => break event.into_owned(),
                Event::Eof => panic!("missing pivot start"),
                _ => {}
            }
            buf.clear();
        };

        let err = parse_ods_pivot_table_full(&mut reader, &start, 1)
            .expect_err("truncated pivot must fail");
        assert!(matches!(err, ParseError::Xml { file, .. } if file == "content.xml"));
    }

    #[test]
    fn parse_ods_pivots_from_xml_rejects_missing_root_end() {
        let xml = br#"<office:document-content
            xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">"#;
        let mut store = IrStore::new();

        let err = parse_ods_pivots_from_xml(xml, &mut store)
            .expect_err("truncated content XML must fail");
        assert!(matches!(err, ParseError::Xml { file, .. } if file == "content.xml"));
    }
}
