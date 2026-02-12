//! XLSX table and pivot parsing.

use crate::error::ParseError;
use docir_core::ir::{PivotCacheRecords, PivotTable, TableColumn, TableDefinition};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::events::Event;
use quick_xml::Reader;

pub(crate) fn parse_table_definition(
    xml: &str,
    table_path: &str,
) -> Result<TableDefinition, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut table = TableDefinition {
        id: NodeId::new(),
        name: None,
        display_name: None,
        ref_range: None,
        header_row_count: None,
        totals_row_count: None,
        columns: Vec::new(),
        span: Some(SourceSpan::new(table_path)),
    };

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => {
                                table.name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"displayName" => {
                                table.display_name =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"ref" => {
                                table.ref_range =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"headerRowCount" => {
                                table.header_row_count =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"totalsRowCount" => {
                                table.totals_row_count =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            _ => {}
                        }
                    }
                }
                b"tableColumn" => {
                    let mut id = None;
                    let mut name = None;
                    let mut totals_row_label = None;
                    let mut totals_row_function = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"id" => id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
                            b"name" => {
                                name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"totalsRowLabel" => {
                                totals_row_label =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"totalsRowFunction" => {
                                totals_row_function =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            _ => {}
                        }
                    }
                    if let Some(id) = id {
                        table.columns.push(TableColumn {
                            id,
                            name,
                            totals_row_label,
                            totals_row_function,
                        });
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) if e.name().as_ref() == b"tableColumn" => {
                let mut id = None;
                let mut name = None;
                let mut totals_row_label = None;
                let mut totals_row_function = None;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"id" => id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
                        b"name" => name = Some(String::from_utf8_lossy(&attr.value).to_string()),
                        b"totalsRowLabel" => {
                            totals_row_label =
                                Some(String::from_utf8_lossy(&attr.value).to_string())
                        }
                        b"totalsRowFunction" => {
                            totals_row_function =
                                Some(String::from_utf8_lossy(&attr.value).to_string())
                        }
                        _ => {}
                    }
                }
                if let Some(id) = id {
                    table.columns.push(TableColumn {
                        id,
                        name,
                        totals_row_label,
                        totals_row_function,
                    });
                }
            }
            Ok(Event::End(e)) if e.name().as_ref() == b"table" => break,
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: table_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(table)
}

pub(crate) fn parse_pivot_table_definition(
    xml: &str,
    pivot_path: &str,
) -> Result<PivotTable, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut pivot = PivotTable {
        id: NodeId::new(),
        name: None,
        cache_id: None,
        ref_range: None,
        span: Some(SourceSpan::new(pivot_path)),
    };

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"pivotTableDefinition" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => {
                                pivot.name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"cacheId" => {
                                pivot.cache_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            _ => {}
                        }
                    }
                }
                b"location" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"ref" {
                            pivot.ref_range =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) if e.name().as_ref() == b"location" => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"ref" {
                        pivot.ref_range = Some(String::from_utf8_lossy(&attr.value).to_string());
                    }
                }
            }
            Ok(Event::End(e)) if e.name().as_ref() == b"pivotTableDefinition" => break,
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: pivot_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(pivot)
}

pub(crate) fn parse_pivot_cache_records(
    xml: &str,
    records_path: &str,
) -> Result<PivotCacheRecords, ParseError> {
    let mut records = PivotCacheRecords::new();
    records.span = Some(SourceSpan::new(records_path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_record = false;
    let mut current_fields: u32 = 0;
    let mut max_fields: u32 = 0;
    let mut counted_records: u32 = 0;
    let mut has_count_attr = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"pivotCacheRecords" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"count" {
                            records.record_count =
                                String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            has_count_attr = records.record_count.is_some();
                        }
                    }
                } else if e.name().as_ref() == b"r" {
                    in_record = true;
                    current_fields = 0;
                } else if in_record {
                    current_fields = current_fields.saturating_add(1);
                }
            }
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"pivotCacheRecords" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"count" {
                            records.record_count =
                                String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            has_count_attr = records.record_count.is_some();
                        }
                    }
                } else if e.name().as_ref() == b"r" {
                    counted_records = counted_records.saturating_add(1);
                    if max_fields < current_fields {
                        max_fields = current_fields;
                    }
                    in_record = false;
                    current_fields = 0;
                } else if in_record {
                    current_fields = current_fields.saturating_add(1);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"r" {
                    counted_records = counted_records.saturating_add(1);
                    if max_fields < current_fields {
                        max_fields = current_fields;
                    }
                    in_record = false;
                    current_fields = 0;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: records_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    if !has_count_attr {
        records.record_count = Some(counted_records);
    }
    if max_fields > 0 {
        records.field_count = Some(max_fields);
    }

    Ok(records)
}
