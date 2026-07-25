//! XLSX table and pivot parsing.

use crate::error::ParseError;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::{XmlScanControl, local_name, scan_xml_events, xml_error};
use docir_core::ir::{PivotCacheRecords, PivotTable, TableColumn, TableDefinition};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::Reader;
use quick_xml::events::Event;
use quick_xml::events::attributes::Attribute;
use std::fmt::Display;
use std::str::FromStr;

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
    scan_xml_events(&mut reader, &mut buf, table_path, |event| {
        match event {
            Event::Start(e) => handle_table_start(&e, &mut table, table_path)?,
            Event::Empty(e) if local_name(e.name().as_ref()) == b"tableColumn" => {
                push_table_column(&e, &mut table.columns, table_path)?;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"table" => {
                return Ok(XmlScanControl::Break);
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok(table)
}

fn handle_table_start(
    e: &quick_xml::events::BytesStart<'_>,
    table: &mut TableDefinition,
    table_path: &str,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"table" => apply_table_attrs(e, table, table_path)?,
        b"tableColumn" => push_table_column(e, &mut table.columns, table_path)?,
        _ => {}
    }
    Ok(())
}

fn apply_table_attrs(
    e: &quick_xml::events::BytesStart<'_>,
    table: &mut TableDefinition,
    table_path: &str,
) -> Result<(), ParseError> {
    visit_table_attributes(e, table_path, |attr| match attr.key.as_ref() {
        b"name" => {
            table.name = Some(lossy_attr_value(attr).to_string());
            Ok(())
        }
        b"displayName" => {
            table.display_name = Some(lossy_attr_value(attr).to_string());
            Ok(())
        }
        b"ref" => {
            table.ref_range = Some(lossy_attr_value(attr).to_string());
            Ok(())
        }
        b"headerRowCount" => {
            table.header_row_count = Some(parse_table_attr(attr, table_path)?);
            Ok(())
        }
        b"totalsRowCount" => {
            table.totals_row_count = Some(parse_table_attr(attr, table_path)?);
            Ok(())
        }
        _ => Ok(()),
    })
}

fn push_table_column(
    e: &quick_xml::events::BytesStart<'_>,
    columns: &mut Vec<TableColumn>,
    table_path: &str,
) -> Result<(), ParseError> {
    let mut id = None;
    let mut name = None;
    let mut totals_row_label = None;
    let mut totals_row_function = None;
    visit_table_attributes(e, table_path, |attr| match attr.key.as_ref() {
        b"id" => {
            id = Some(parse_table_attr(attr, table_path)?);
            Ok(())
        }
        b"name" => {
            name = Some(lossy_attr_value(attr).to_string());
            Ok(())
        }
        b"totalsRowLabel" => {
            totals_row_label = Some(lossy_attr_value(attr).to_string());
            Ok(())
        }
        b"totalsRowFunction" => {
            totals_row_function = Some(lossy_attr_value(attr).to_string());
            Ok(())
        }
        _ => Ok(()),
    })?;
    if let Some(id) = id {
        columns.push(TableColumn {
            id,
            name,
            totals_row_label,
            totals_row_function,
        });
    }
    Ok(())
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
    scan_xml_events(&mut reader, &mut buf, pivot_path, |event| {
        match event {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"pivotTableDefinition" => {
                    visit_table_attributes(&e, pivot_path, |attr| match attr.key.as_ref() {
                        b"name" => {
                            pivot.name = Some(lossy_attr_value(attr).to_string());
                            Ok(())
                        }
                        b"cacheId" => {
                            pivot.cache_id = Some(parse_table_attr(attr, pivot_path)?);
                            Ok(())
                        }
                        _ => Ok(()),
                    })?;
                }
                b"location" => {
                    visit_table_attributes(&e, pivot_path, |attr| {
                        if attr.key.as_ref() == b"ref" {
                            pivot.ref_range = Some(lossy_attr_value(attr).to_string());
                        }
                        Ok(())
                    })?;
                }
                _ => {}
            },
            Event::Empty(e) if local_name(e.name().as_ref()) == b"location" => {
                visit_table_attributes(&e, pivot_path, |attr| {
                    if attr.key.as_ref() == b"ref" {
                        pivot.ref_range = Some(lossy_attr_value(attr).to_string());
                    }
                    Ok(())
                })?;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"pivotTableDefinition" => {
                return Ok(XmlScanControl::Break);
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

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

    scan_xml_events(&mut reader, &mut buf, records_path, |event| {
        match event {
            Event::Start(e) => {
                if local_name(e.name().as_ref()) == b"pivotCacheRecords" {
                    visit_table_attributes(&e, records_path, |attr| {
                        if attr.key.as_ref() == b"count" {
                            records.record_count = Some(parse_table_attr(attr, records_path)?);
                            has_count_attr = true;
                        }
                        Ok(())
                    })?;
                } else if local_name(e.name().as_ref()) == b"r" {
                    in_record = true;
                    current_fields = 0;
                } else if in_record {
                    current_fields = current_fields.saturating_add(1);
                }
            }
            Event::Empty(e) => {
                if local_name(e.name().as_ref()) == b"pivotCacheRecords" {
                    visit_table_attributes(&e, records_path, |attr| {
                        if attr.key.as_ref() == b"count" {
                            records.record_count = Some(parse_table_attr(attr, records_path)?);
                            has_count_attr = true;
                        }
                        Ok(())
                    })?;
                } else if local_name(e.name().as_ref()) == b"r" {
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
            Event::End(e) if local_name(e.name().as_ref()) == b"r" => {
                counted_records = counted_records.saturating_add(1);
                if max_fields < current_fields {
                    max_fields = current_fields;
                }
                in_record = false;
                current_fields = 0;
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    if !has_count_attr {
        records.record_count = Some(counted_records);
    }
    if max_fields > 0 {
        records.field_count = Some(max_fields);
    }

    Ok(records)
}

fn visit_table_attributes<F>(
    start: &quick_xml::events::BytesStart<'_>,
    path: &str,
    mut visit: F,
) -> Result<(), ParseError>
where
    F: FnMut(&Attribute<'_>) -> Result<(), ParseError>,
{
    for attr in start.attributes() {
        let attr = attr.map_err(|err| xml_error(path, err))?;
        visit(&attr)?;
    }
    Ok(())
}

fn parse_table_attr<T>(attr: &Attribute<'_>, path: &str) -> Result<T, ParseError>
where
    T: FromStr,
    T::Err: Display,
{
    lossy_attr_value(attr)
        .parse::<T>()
        .map_err(|err| xml_error(path, err))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_table_definition_reads_table_attrs_and_columns() {
        let xml = r#"
            <table name="SalesTable" displayName="Sales" ref="A1:C10" headerRowCount="1" totalsRowCount="1">
              <tableColumns count="2">
                <tableColumn id="1" name="Amount" totalsRowFunction="sum"></tableColumn>
                <tableColumn id="2" name="Region" totalsRowLabel="Total"/>
                <tableColumn name="IgnoredWithoutId"/>
              </tableColumns>
            </table>
        "#;

        let parsed =
            parse_table_definition(xml, "xl/tables/table1.xml").expect("table should parse");
        assert_eq!(parsed.name.as_deref(), Some("SalesTable"));
        assert_eq!(parsed.display_name.as_deref(), Some("Sales"));
        assert_eq!(parsed.ref_range.as_deref(), Some("A1:C10"));
        assert_eq!(parsed.header_row_count, Some(1));
        assert_eq!(parsed.totals_row_count, Some(1));
        assert_eq!(parsed.columns.len(), 2);
        assert_eq!(parsed.columns[0].id, 1);
        assert_eq!(parsed.columns[0].name.as_deref(), Some("Amount"));
        assert_eq!(
            parsed.columns[0].totals_row_function.as_deref(),
            Some("sum")
        );
        assert_eq!(parsed.columns[1].id, 2);
        assert_eq!(parsed.columns[1].totals_row_label.as_deref(), Some("Total"));
        assert!(parsed.span.is_some());
    }

    #[test]
    fn parse_table_definition_returns_xml_error_for_malformed_input() {
        let xml = "<table><tableColumns><tableColumn id='1' name='X'></table>";
        let err = parse_table_definition(xml, "xl/tables/broken.xml")
            .expect_err("malformed table xml should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/tables/broken.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_table_definition_reports_malformed_attributes() {
        let xml = r#"
            <table name="SalesTable" name="Other" ref="A1:C10">
              <tableColumns count="1">
                <tableColumn id="1" name="Amount"/>
              </tableColumns>
            </table>
        "#;
        let err = parse_table_definition(xml, "xl/tables/broken-attrs.xml")
            .expect_err("malformed table attributes should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/tables/broken-attrs.xml"),
            other => panic!("unexpected error: {other:?}"),
        }

        let xml = r#"
            <table name="SalesTable" ref="A1:C10" headerRowCount="bad">
              <tableColumns count="1">
                <tableColumn id="1" name="Amount"/>
              </tableColumns>
            </table>
        "#;
        let err = parse_table_definition(xml, "xl/tables/broken-count.xml")
            .expect_err("malformed table counts should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/tables/broken-count.xml"),
            other => panic!("unexpected error: {other:?}"),
        }

        let xml = r#"
            <table name="SalesTable" ref="A1:C10">
              <tableColumns count="1">
                <tableColumn id="1" id="2" name="Amount"/>
              </tableColumns>
            </table>
        "#;
        let err = parse_table_definition(xml, "xl/tables/broken-column-attrs.xml")
            .expect_err("malformed table column attributes should fail");
        match err {
            ParseError::Xml { file, .. } => {
                assert_eq!(file, "xl/tables/broken-column-attrs.xml")
            }
            other => panic!("unexpected error: {other:?}"),
        }

        let xml = r#"
            <table name="SalesTable" ref="A1:C10">
              <tableColumns count="1">
                <tableColumn id="bad" name="Amount"/>
              </tableColumns>
            </table>
        "#;
        let err = parse_table_definition(xml, "xl/tables/broken-column-id.xml")
            .expect_err("malformed table column id should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/tables/broken-column-id.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_pivot_table_definition_reads_name_cache_and_location() {
        let xml = r#"
            <pivotTableDefinition name="PivotA" cacheId="7">
              <location ref="B3:E20"/>
            </pivotTableDefinition>
        "#;
        let parsed = parse_pivot_table_definition(xml, "xl/pivotTables/pivotTable1.xml")
            .expect("pivot should parse");
        assert_eq!(parsed.name.as_deref(), Some("PivotA"));
        assert_eq!(parsed.cache_id, Some(7));
        assert_eq!(parsed.ref_range.as_deref(), Some("B3:E20"));
        assert!(parsed.span.is_some());
    }

    #[test]
    fn parse_pivot_table_definition_handles_empty_location_tag() {
        let xml = r#"
            <pivotTableDefinition name="PivotB" cacheId="9">
              <location ref="A1:A5"/>
            </pivotTableDefinition>
        "#;
        let parsed = parse_pivot_table_definition(xml, "xl/pivotTables/pivotTable2.xml")
            .expect("pivot should parse");
        assert_eq!(parsed.ref_range.as_deref(), Some("A1:A5"));
    }

    #[test]
    fn parse_pivot_table_definition_reports_malformed_attributes() {
        let xml = r#"
            <pivotTableDefinition name="PivotA" name="Other" cacheId="7">
              <location ref="B3:E20"/>
            </pivotTableDefinition>
        "#;
        let err = parse_pivot_table_definition(xml, "xl/pivotTables/broken-attrs.xml")
            .expect_err("malformed pivot attributes should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/pivotTables/broken-attrs.xml"),
            other => panic!("unexpected error: {other:?}"),
        }

        let xml = r#"
            <pivotTableDefinition name="PivotA" cacheId="7">
              <location ref="B3:E20" ref="B3:F20"/>
            </pivotTableDefinition>
        "#;
        let err = parse_pivot_table_definition(xml, "xl/pivotTables/broken-location-attrs.xml")
            .expect_err("malformed pivot location attributes should fail");
        match err {
            ParseError::Xml { file, .. } => {
                assert_eq!(file, "xl/pivotTables/broken-location-attrs.xml")
            }
            other => panic!("unexpected error: {other:?}"),
        }

        let xml = r#"
            <pivotTableDefinition name="PivotA" cacheId="bad">
              <location ref="B3:E20"/>
            </pivotTableDefinition>
        "#;
        let err = parse_pivot_table_definition(xml, "xl/pivotTables/broken-cache-id.xml")
            .expect_err("malformed pivot cache id should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/pivotTables/broken-cache-id.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_pivot_cache_records_uses_count_attribute_when_present() {
        let xml = r#"
            <pivotCacheRecords count="4">
              <r><x v="1"/><x v="2"/></r>
              <r><x v="3"/></r>
            </pivotCacheRecords>
        "#;
        let parsed = parse_pivot_cache_records(xml, "xl/pivotCache/pivotCacheRecords1.xml")
            .expect("records should parse");
        assert_eq!(parsed.record_count, Some(4));
        assert_eq!(parsed.field_count, Some(2));
        assert!(parsed.span.is_some());
    }

    #[test]
    fn parse_pivot_cache_records_counts_records_when_count_missing() {
        let xml = r#"
            <pivotCacheRecords>
              <r><x v="1"/></r>
              <r><x v="2"/><x v="3"/><x v="4"/></r>
              <r/>
            </pivotCacheRecords>
        "#;
        let parsed = parse_pivot_cache_records(xml, "xl/pivotCache/pivotCacheRecords2.xml")
            .expect("records should parse");
        assert_eq!(parsed.record_count, Some(3));
        assert_eq!(parsed.field_count, Some(3));
    }

    #[test]
    fn parse_pivot_cache_records_reports_malformed_attributes() {
        let xml = r#"
            <pivotCacheRecords count="4" count="5">
              <r><x v="1"/></r>
            </pivotCacheRecords>
        "#;
        let err = parse_pivot_cache_records(xml, "xl/pivotCache/broken-attrs.xml")
            .expect_err("malformed cache record attributes should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/pivotCache/broken-attrs.xml"),
            other => panic!("unexpected error: {other:?}"),
        }

        let xml = r#"
            <pivotCacheRecords count="bad">
              <r><x v="1"/></r>
            </pivotCacheRecords>
        "#;
        let err = parse_pivot_cache_records(xml, "xl/pivotCache/broken-count.xml")
            .expect_err("malformed cache record count should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/pivotCache/broken-count.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_pivot_cache_records_returns_xml_error_for_malformed_input() {
        let xml = "<pivotCacheRecords><r><x v='1'></pivotCacheRecords";
        let err = parse_pivot_cache_records(xml, "xl/pivotCache/broken.xml")
            .expect_err("malformed cache records xml should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/pivotCache/broken.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }
}
