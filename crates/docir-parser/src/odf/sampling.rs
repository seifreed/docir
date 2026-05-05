use super::helpers::{parse_ods_covered_cell, parse_ods_covered_cell_empty, OdsRow};
use super::ods::{parse_ods_cell, parse_ods_cell_empty};
use super::{spreadsheet, OdfReader};
use crate::error::ParseError;
use crate::xml_utils::{attr_value_by_suffix, local_name, xml_error};
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
    let mut cells = Vec::new();
    let mut col_idx: u32 = 0;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"table-cell" => {
                    if col_idx >= sample_cols {
                        let repeat = attr_value_by_suffix(&e, &[b":number-columns-repeated"])
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        col_idx = col_idx.saturating_add(repeat);
                        spreadsheet::skip_element(reader, e.name().as_ref())?;
                    } else {
                        let cell = parse_ods_cell(reader, &e, store, style_map, next_style_id)?;
                        col_idx = col_idx.saturating_add(cell.col_repeat);
                        cells.push(cell);
                    }
                }
                b"covered-table-cell" => {
                    if col_idx >= sample_cols {
                        let repeat = attr_value_by_suffix(&e, &[b":number-columns-repeated"])
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        col_idx = col_idx.saturating_add(repeat);
                        spreadsheet::skip_element(reader, e.name().as_ref())?;
                    } else {
                        let cell = parse_ods_covered_cell(reader, &e)?;
                        col_idx = col_idx.saturating_add(cell.col_repeat);
                        cells.push(cell);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                b"table-cell" => {
                    if col_idx < sample_cols {
                        let cell = parse_ods_cell_empty(&e, style_map, next_style_id)?;
                        col_idx = col_idx.saturating_add(cell.col_repeat);
                        cells.push(cell);
                    } else {
                        let repeat = attr_value_by_suffix(&e, &[b":number-columns-repeated"])
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        col_idx = col_idx.saturating_add(repeat);
                    }
                }
                b"covered-table-cell" => {
                    if col_idx < sample_cols {
                        let cell = parse_ods_covered_cell_empty(&e)?;
                        col_idx = col_idx.saturating_add(cell.col_repeat);
                        cells.push(cell);
                    } else {
                        let repeat = attr_value_by_suffix(&e, &[b":number-columns-repeated"])
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        col_idx = col_idx.saturating_add(repeat);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"table-row" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error("content.xml", e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(OdsRow { cells })
}
