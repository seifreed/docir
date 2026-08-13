use crate::error::ParseError;
use crate::odf::{
    OdfReader,
    ods::{parse_ods_cell, parse_ods_cell_empty},
};
use crate::xml_utils::{
    XmlScanControl, is_end_event_local, local_name, scan_xml_events_until_end,
    try_attr_value_by_suffix, xml_error,
};
use docir_core::ir::{CellFormula, CellValue, MergedCellRange};
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;

#[derive(Debug, Clone)]
pub(crate) struct OdsRow {
    pub(crate) cells: Vec<OdsCellData>,
}

#[derive(Debug, Clone)]
pub(crate) struct OdsCellData {
    pub(crate) value: CellValue,
    pub(crate) formula: Option<CellFormula>,
    pub(crate) style_id: Option<u32>,
    pub(crate) col_repeat: u32,
    pub(crate) validation_name: Option<String>,
    pub(crate) col_span: Option<u32>,
    pub(crate) row_span: Option<u32>,
    pub(crate) is_covered: bool,
}

impl OdsCellData {
    pub(crate) fn should_emit(&self) -> bool {
        !matches!(self.value, CellValue::Empty) || self.formula.is_some()
    }

    pub(crate) fn merge_range(
        &self,
        row: u32,
        col: u32,
    ) -> Result<Option<MergedCellRange>, ParseError> {
        let col_span = self.col_span.unwrap_or(1);
        let row_span = self.row_span.unwrap_or(1);
        if col_span == 0 || row_span == 0 {
            return Err(ParseError::InvalidStructure(
                "ODF merged cell spans must be positive".to_string(),
            ));
        }
        if col_span > 1 || row_span > 1 {
            let end_col = col.checked_add(col_span - 1).ok_or_else(|| {
                ParseError::InvalidStructure("ODF merged cell column overflow".to_string())
            })?;
            let end_row = row.checked_add(row_span - 1).ok_or_else(|| {
                ParseError::InvalidStructure("ODF merged cell row overflow".to_string())
            })?;
            Ok(Some(MergedCellRange {
                start_col: col,
                start_row: row,
                end_col,
                end_row,
            }))
        } else {
            Ok(None)
        }
    }
}

pub(crate) fn parse_ods_row(
    reader: &mut OdfReader<'_>,
    _start: &BytesStart<'_>,
    store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<OdsRow, ParseError> {
    let mut state = OdsRowParseState::default();
    let mut buf = Vec::new();
    scan_xml_events_until_end(
        reader,
        &mut buf,
        "content.xml",
        |event| is_end_event_local(event, b"table-row"),
        |reader, event| {
            dispatch_ods_row_event(reader, event, &mut state, store, style_map, next_style_id)?;
            Ok(XmlScanControl::Continue)
        },
    )?;

    Ok(OdsRow { cells: state.cells })
}

#[derive(Default)]
struct OdsRowParseState {
    cells: Vec<OdsCellData>,
}

fn dispatch_ods_row_event(
    reader: &mut OdfReader<'_>,
    event: &Event<'_>,
    state: &mut OdsRowParseState,
    store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<(), ParseError> {
    match event {
        Event::Start(e) => {
            handle_ods_row_start_event(reader, e, store, style_map, next_style_id, state)
        }
        Event::Empty(e) => handle_ods_row_empty_event(e, style_map, next_style_id, state),
        _ => Ok(()),
    }
}

fn handle_ods_row_start_event(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
    state: &mut OdsRowParseState,
) -> Result<(), ParseError> {
    match local_name(start.name().as_ref()) {
        b"table-cell" => {
            let cell = parse_ods_cell(reader, start, store, style_map, next_style_id)?;
            state.cells.push(cell);
        }
        b"covered-table-cell" => {
            let cell = parse_ods_covered_cell(reader, start)?;
            state.cells.push(cell);
        }
        _ => {}
    }
    Ok(())
}

fn handle_ods_row_empty_event(
    start: &BytesStart<'_>,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
    state: &mut OdsRowParseState,
) -> Result<(), ParseError> {
    match local_name(start.name().as_ref()) {
        b"table-cell" => {
            let cell = parse_ods_cell_empty(start, style_map, next_style_id)?;
            state.cells.push(cell);
        }
        b"covered-table-cell" => {
            let cell = parse_ods_covered_cell_empty(start)?;
            state.cells.push(cell);
        }
        _ => {}
    }
    Ok(())
}

pub(crate) fn parse_ods_covered_cell(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
) -> Result<OdsCellData, ParseError> {
    let cell = covered_cell_from_start(start)?;
    let mut buf = Vec::new();
    scan_xml_events_until_end(
        reader,
        &mut buf,
        "content.xml",
        |event| is_end_event_local(event, b"covered-table-cell"),
        |_reader, _event| Ok(XmlScanControl::Continue),
    )?;
    Ok(cell)
}

pub(crate) fn parse_ods_covered_cell_empty(
    start: &BytesStart<'_>,
) -> Result<OdsCellData, ParseError> {
    covered_cell_from_start(start)
}

fn covered_cell_from_start(start: &BytesStart<'_>) -> Result<OdsCellData, ParseError> {
    let col_repeat =
        try_attr_value_by_suffix(start, &[b":number-columns-repeated"], "content.xml")?
            .map(|v| parse_u32_attr(&v))
            .transpose()?
            .unwrap_or(1);
    let col_span = try_attr_value_by_suffix(start, &[b":number-columns-spanned"], "content.xml")?
        .map(|v| parse_u32_attr(&v))
        .transpose()?;
    let row_span = try_attr_value_by_suffix(start, &[b":number-rows-spanned"], "content.xml")?
        .map(|v| parse_u32_attr(&v))
        .transpose()?;
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

fn parse_u32_attr(value: &str) -> Result<u32, ParseError> {
    let parsed = value.parse().map_err(|err| xml_error("content.xml", err))?;
    if parsed == 0 {
        return Err(xml_error(
            "content.xml",
            "ODF cell span/repeat attributes must be positive",
        ));
    }
    Ok(parsed)
}

pub(crate) fn column_index_to_name(mut index: u32) -> Result<String, ParseError> {
    let mut name = String::new();
    index = index
        .checked_add(1)
        .ok_or_else(|| ParseError::InvalidStructure("ODF column index overflow".to_string()))?;
    while index > 0 {
        let rem = ((index - 1) % 26) as u8;
        name.push((b'A' + rem) as char);
        index = (index - 1) / 26;
    }
    Ok(name.chars().rev().collect())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn covered_cell_empty_maps_repeat_and_spans() {
        let mut reader = quick_xml::Reader::from_str(
            r#"<table:covered-table-cell table:number-columns-repeated="2" table:number-columns-spanned="3" table:number-rows-spanned="4" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/>"#,
        );
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let start = match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => e.into_owned(),
            _ => panic!("expected covered-table-cell"),
        };

        let cell = parse_ods_covered_cell_empty(&start).expect("covered cell parse");
        assert_eq!(cell.col_repeat, 2);
        assert_eq!(cell.col_span, Some(3));
        assert_eq!(cell.row_span, Some(4));
        assert!(cell.is_covered);
    }

    #[test]
    fn covered_cell_empty_reports_malformed_span_attributes() {
        let mut reader = quick_xml::Reader::from_str(
            r#"<table:covered-table-cell table:number-columns-spanned="2" table:number-columns-spanned="3" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/>"#,
        );
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let start = match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => e.into_owned(),
            _ => panic!("expected covered-table-cell"),
        };

        let err = parse_ods_covered_cell_empty(&start).expect_err("malformed span must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
            other => panic!("expected xml error, got {other:?}"),
        }
    }

    #[test]
    fn covered_cell_empty_reports_malformed_numeric_attributes() {
        for xml in [
            r#"<table:covered-table-cell table:number-columns-repeated="bad" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/>"#,
            r#"<table:covered-table-cell table:number-columns-spanned="bad" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/>"#,
            r#"<table:covered-table-cell table:number-rows-spanned="bad" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/>"#,
        ] {
            let mut reader = quick_xml::Reader::from_str(xml);
            reader.config_mut().trim_text(true);
            let mut buf = Vec::new();
            let start = match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) => e.into_owned(),
                _ => panic!("expected covered-table-cell"),
            };

            let err = parse_ods_covered_cell_empty(&start)
                .expect_err("malformed covered cell numeric attributes must fail");
            match err {
                ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
                other => panic!("expected xml error, got {other:?}"),
            }
        }
    }

    #[test]
    fn column_index_to_name_handles_edges() {
        assert_eq!(column_index_to_name(0).expect("column name"), "A");
        assert_eq!(column_index_to_name(25).expect("column name"), "Z");
        assert_eq!(column_index_to_name(26).expect("column name"), "AA");
        assert_eq!(column_index_to_name(701).expect("column name"), "ZZ");
    }
}
