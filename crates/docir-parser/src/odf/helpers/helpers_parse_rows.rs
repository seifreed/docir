use crate::error::ParseError;
use crate::odf::{
    ods::{parse_ods_cell, parse_ods_cell_empty},
    OdfReader,
};
use crate::xml_utils::{
    attr_value_by_suffix, is_end_event_local, local_name, scan_xml_events_until_end, XmlScanControl,
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

    pub(crate) fn merge_range(&self, row: u32, col: u32) -> Option<MergedCellRange> {
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
    let cell = covered_cell_from_start(start);
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
    Ok(covered_cell_from_start(start))
}

fn covered_cell_from_start(start: &BytesStart<'_>) -> OdsCellData {
    let col_repeat = attr_value_by_suffix(start, &[b":number-columns-repeated"])
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let col_span = attr_value_by_suffix(start, &[b":number-columns-spanned"])
        .and_then(|v| v.parse::<u32>().ok());
    let row_span =
        attr_value_by_suffix(start, &[b":number-rows-spanned"]).and_then(|v| v.parse::<u32>().ok());
    OdsCellData {
        value: CellValue::Empty,
        formula: None,
        style_id: None,
        col_repeat,
        validation_name: None,
        col_span,
        row_span,
        is_covered: true,
    }
}

pub(crate) fn column_index_to_name(mut index: u32) -> String {
    let mut name = String::new();
    index += 1;
    while index > 0 {
        let rem = ((index - 1) % 26) as u8;
        name.push((b'A' + rem) as char);
        index = (index - 1) / 26;
    }
    name.chars().rev().collect()
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
    fn column_index_to_name_handles_edges() {
        assert_eq!(column_index_to_name(0), "A");
        assert_eq!(column_index_to_name(25), "Z");
        assert_eq!(column_index_to_name(26), "AA");
        assert_eq!(column_index_to_name(701), "ZZ");
    }
}
