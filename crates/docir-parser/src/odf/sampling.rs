use super::helpers::{OdsCellData, OdsRow, parse_ods_covered_cell, parse_ods_covered_cell_empty};
use super::ods::{parse_ods_cell, parse_ods_cell_empty};
use super::{OdfReader, spreadsheet};
use crate::error::ParseError;
use crate::xml_utils::{local_name, try_attr_value_by_suffix, xml_error};
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;

pub(super) fn parse_ods_row_sample(
    reader: &mut OdfReader<'_>,
    _start: &BytesStart<'_>,
    store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
    sample_cols: u32,
) -> Result<OdsRow, ParseError> {
    let mut buf = Vec::new();
    let mut state = OdsSampleState::new(sample_cols);
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"table-cell" => handle_sample_table_cell(
                    reader,
                    &e,
                    &mut state,
                    store,
                    style_map,
                    next_style_id,
                )?,
                b"covered-table-cell" => handle_sample_covered_cell(reader, &e, &mut state)?,
                _ => {}
            },
            Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                b"table-cell" => {
                    handle_empty_sample_table_cell(&e, &mut state, style_map, next_style_id)?
                }
                b"covered-table-cell" => {
                    handle_empty_sample_covered_cell(&e, &mut state)?;
                }
                _ => {}
            },
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"table-row" => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error("content.xml", e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(OdsRow { cells: state.cells })
}

struct OdsSampleState {
    cells: Vec<OdsCellData>,
    col_idx: u32,
    sample_cols: u32,
}

impl OdsSampleState {
    fn new(sample_cols: u32) -> Self {
        Self {
            cells: Vec::new(),
            col_idx: 0,
            sample_cols,
        }
    }

    fn push_cell(&mut self, cell: OdsCellData) {
        self.col_idx = self.col_idx.saturating_add(cell.col_repeat);
        self.cells.push(cell);
    }

    fn skip_columns(&mut self, repeat: u32) {
        self.col_idx = self.col_idx.saturating_add(repeat);
    }
}

fn handle_sample_table_cell(
    reader: &mut OdfReader<'_>,
    e: &BytesStart<'_>,
    state: &mut OdsSampleState,
    store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<(), ParseError> {
    if state.col_idx >= state.sample_cols {
        skip_sampled_cell(reader, e, state)
    } else {
        let cell = parse_ods_cell(reader, e, store, style_map, next_style_id)?;
        state.push_cell(cell);
        Ok(())
    }
}

fn handle_sample_covered_cell(
    reader: &mut OdfReader<'_>,
    e: &BytesStart<'_>,
    state: &mut OdsSampleState,
) -> Result<(), ParseError> {
    if state.col_idx >= state.sample_cols {
        skip_sampled_cell(reader, e, state)
    } else {
        let cell = parse_ods_covered_cell(reader, e)?;
        state.push_cell(cell);
        Ok(())
    }
}

fn handle_empty_sample_table_cell(
    e: &BytesStart<'_>,
    state: &mut OdsSampleState,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<(), ParseError> {
    if state.col_idx < state.sample_cols {
        let cell = parse_ods_cell_empty(e, style_map, next_style_id)?;
        state.push_cell(cell);
    } else {
        state.skip_columns(repeated_columns(e)?);
    }
    Ok(())
}

fn handle_empty_sample_covered_cell(
    e: &BytesStart<'_>,
    state: &mut OdsSampleState,
) -> Result<(), ParseError> {
    if state.col_idx < state.sample_cols {
        let cell = parse_ods_covered_cell_empty(e)?;
        state.push_cell(cell);
    } else {
        state.skip_columns(repeated_columns(e)?);
    }
    Ok(())
}

fn skip_sampled_cell(
    reader: &mut OdfReader<'_>,
    e: &BytesStart<'_>,
    state: &mut OdsSampleState,
) -> Result<(), ParseError> {
    state.skip_columns(repeated_columns(e)?);
    spreadsheet::skip_element(reader, e.name().as_ref())
}

fn repeated_columns(e: &BytesStart<'_>) -> Result<u32, ParseError> {
    let value = try_attr_value_by_suffix(e, &[b":number-columns-repeated"], "content.xml")?;
    value
        .map(|value| {
            let parsed = value
                .parse::<u32>()
                .map_err(|err| xml_error("content.xml", err))?;
            if parsed == 0 {
                return Err(xml_error(
                    "content.xml",
                    "table:number-columns-repeated must be positive",
                ));
            }
            Ok(parsed)
        })
        .transpose()
        .map(|value| value.unwrap_or(1))
}
