use crate::ooxml::part_utils::read_relationships;
use crate::ooxml::relationships::Relationships;
use crate::ooxml::xlsx::{
    parse_pivot_cache_records, rel_type, IRNode, ParseError, PivotCache, XlsxParser,
};
use crate::xml_utils::{reader_from_str, scan_xml_events_until_end, XmlScanControl};
use crate::zip_handler::PackageReader;
use docir_core::types::SourceSpan;
use quick_xml::events::Event;

pub(super) fn parse_pivot_cache_impl(
    parser: &mut XlsxParser,
    zip: &mut impl PackageReader,
    xml: &str,
    cache_path: &str,
    cache_id: u32,
) -> Result<PivotCache, ParseError> {
    let mut reader = reader_from_str(xml);

    let mut cache = PivotCache::new(cache_id);
    cache.span = Some(SourceSpan::new(cache_path));

    let mut buf = Vec::new();
    scan_xml_events_until_end(
        &mut reader,
        &mut buf,
        cache_path,
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"pivotCacheDefinition"),
        |_reader, event| {
            if let Event::Start(e) | Event::Empty(e) = event {
                match e.name().as_ref() {
                    b"cacheSource" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"type" {
                                cache.cache_source =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            if attr.key.as_ref() == b"connectionId" {
                                let conn = String::from_utf8_lossy(&attr.value).to_string();
                                cache.cache_source = Some(format!("connection:{conn}"));
                            }
                        }
                    }
                    b"worksheetSource" => {
                        let mut sheet = None;
                        let mut range = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"sheet" => {
                                    sheet = Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"ref" => {
                                    range = Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                _ => {}
                            }
                        }
                        if let Some(sheet) = sheet {
                            let range = range.unwrap_or_else(|| "-".to_string());
                            cache.cache_source = Some(format!("worksheet:{sheet}!{range}"));
                        }
                    }
                    _ => {}
                }
            }
            Ok(XmlScanControl::Continue)
        },
    )?;

    attach_pivot_records(parser, zip, cache_path, cache_id, &mut cache)?;
    Ok(cache)
}

fn attach_pivot_records(
    parser: &mut XlsxParser,
    zip: &mut impl PackageReader,
    cache_path: &str,
    cache_id: u32,
    cache: &mut PivotCache,
) -> Result<(), ParseError> {
    let rels = read_relationships(zip, cache_path)?;
    if let Some(rel) = rels.get_first_by_type(rel_type::PIVOT_CACHE_RECORDS) {
        let records_path = Relationships::resolve_target(cache_path, &rel.target);
        if zip.contains(&records_path) {
            let records_xml = zip.read_file_string(&records_path)?;
            let mut records = parse_pivot_cache_records(&records_xml, &records_path)?;
            records.cache_id = Some(cache_id);
            cache.record_count = records.record_count;
            let rec_id = records.id;
            parser.store.insert(IRNode::PivotCacheRecords(records));
            cache.records = Some(rec_id);
        }
    }
    Ok(())
}
