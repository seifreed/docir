use super::super::helpers::{
    parse_ods_conditional_formatting, parse_ods_conditional_formatting_empty, parse_ods_row,
    OdsCellData, ValidationDef,
};
use super::ods_postprocess::emit_full_row_cells;
use super::{
    attach_shapes_as_drawing, emit_sampled_row_cells, flush_validation_ranges,
    infer_cell_value_type_and_attr, parse_cell_formula, parse_cell_value_empty,
    parse_cell_value_with_text, push_conditional_format, read_ods_cell_text, resolve_style_id,
    row_repeat_from, RowBuildState,
};
use crate::odf::{
    evaluate_ods_formulas, parse_ods_row_sample, spreadsheet, CellValue, IrStore, NodeId,
    OdfLimitCounter, OdfReader, ParseError, Worksheet,
};
use crate::xml_utils::{
    attr_value_by_suffix, dispatch_start_or_empty, is_end_event_local, local_name,
    scan_xml_events_with_reader, XmlScanControl,
};
use quick_xml::events::BytesStart;
use std::collections::HashMap;

struct FullTableContext<'a> {
    store: &'a mut IrStore,
    limits: &'a dyn OdfLimitCounter,
    style_map: &'a mut HashMap<String, u32>,
    next_style_id: &'a mut u32,
    row_idx: &'a mut u32,
    worksheet: &'a mut Worksheet,
    validation_ranges: &'a mut HashMap<String, Vec<String>>,
    cell_values: &'a mut HashMap<(u32, u32), CellValue>,
    formula_cells: &'a mut Vec<(NodeId, u32, u32, String)>,
    formula_map: &'a mut HashMap<(u32, u32), String>,
    shapes: &'a mut Vec<NodeId>,
}

struct FastTableContext<'a> {
    store: &'a mut IrStore,
    limits: &'a dyn OdfLimitCounter,
    style_map: &'a mut HashMap<String, u32>,
    next_style_id: &'a mut u32,
    row_idx: &'a mut u32,
    worksheet: &'a mut Worksheet,
    validation_ranges: &'a mut HashMap<String, Vec<String>>,
}

pub(crate) fn parse_ods_table(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    sheet_id: u32,
    store: &mut IrStore,
    validations: &HashMap<String, ValidationDef>,
    limits: &dyn OdfLimitCounter,
) -> Result<Worksheet, ParseError> {
    if limits.fast_mode() {
        return parse_ods_table_fast(reader, start, sheet_id, store, validations, limits);
    }
    let name =
        attr_value_by_suffix(start, &[b":name"]).unwrap_or_else(|| format!("Sheet{sheet_id}"));
    let mut worksheet = Worksheet::new(name, sheet_id);
    let mut buf = Vec::new();
    let mut row_idx: u32 = 0;
    let mut style_map: HashMap<String, u32> = HashMap::new();
    let mut next_style_id = 1u32;
    let mut validation_ranges: HashMap<String, Vec<String>> = HashMap::new();
    let mut shapes: Vec<NodeId> = Vec::new();
    let mut cell_values: HashMap<(u32, u32), CellValue> = HashMap::new();
    let mut formula_cells: Vec<(NodeId, u32, u32, String)> = Vec::new();
    let mut formula_map: HashMap<(u32, u32), String> = HashMap::new();

    scan_xml_events_with_reader(reader, &mut buf, "content.xml", |reader, event| {
        if is_end_event_local(&event, b"table") {
            return Ok(XmlScanControl::Break);
        }
        let _ = dispatch_start_or_empty(reader, &event, |reader, e, is_start| {
            if is_start {
                handle_table_start_full(
                    reader,
                    e,
                    FullTableContext {
                        store,
                        limits,
                        style_map: &mut style_map,
                        next_style_id: &mut next_style_id,
                        row_idx: &mut row_idx,
                        worksheet: &mut worksheet,
                        validation_ranges: &mut validation_ranges,
                        cell_values: &mut cell_values,
                        formula_cells: &mut formula_cells,
                        formula_map: &mut formula_map,
                        shapes: &mut shapes,
                    },
                )
            } else {
                handle_table_empty_full(
                    reader,
                    e,
                    store,
                    limits,
                    &mut row_idx,
                    &mut worksheet,
                    &mut shapes,
                )
            }
        })?;
        Ok(XmlScanControl::Continue)
    })?;

    if !formula_cells.is_empty() {
        evaluate_ods_formulas(
            &worksheet.name,
            &formula_cells,
            store,
            &mut cell_values,
            &formula_map,
        );
    }

    flush_validation_ranges(validation_ranges, validations, store, &mut worksheet, None);
    attach_shapes_as_drawing(shapes, store, &mut worksheet);

    Ok(worksheet)
}

pub(crate) fn parse_ods_table_fast(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    sheet_id: u32,
    store: &mut IrStore,
    validations: &HashMap<String, ValidationDef>,
    limits: &dyn OdfLimitCounter,
) -> Result<Worksheet, ParseError> {
    let name =
        attr_value_by_suffix(start, &[b":name"]).unwrap_or_else(|| format!("Sheet{sheet_id}"));
    let mut worksheet = Worksheet::new(name, sheet_id);
    let mut buf = Vec::new();
    let mut row_idx: u32 = 0;
    let mut style_map: HashMap<String, u32> = HashMap::new();
    let mut next_style_id = 1u32;
    let mut validation_ranges: HashMap<String, Vec<String>> = HashMap::new();
    let sample_rows = limits.sample_rows();
    let sample_cols = limits.sample_cols();
    let sample_enabled = sample_rows > 0 && sample_cols > 0;

    scan_xml_events_with_reader(reader, &mut buf, "content.xml", |reader, event| {
        if is_end_event_local(&event, b"table") {
            return Ok(XmlScanControl::Break);
        }
        let _ = dispatch_start_or_empty(reader, &event, |reader, e, is_start| {
            if is_start {
                handle_table_start_fast(
                    reader,
                    e,
                    sample_rows,
                    sample_cols,
                    sample_enabled,
                    FastTableContext {
                        store,
                        limits,
                        style_map: &mut style_map,
                        next_style_id: &mut next_style_id,
                        row_idx: &mut row_idx,
                        worksheet: &mut worksheet,
                        validation_ranges: &mut validation_ranges,
                    },
                )
            } else {
                handle_table_empty_fast(e, limits, &mut row_idx)
            }
        })?;
        Ok(XmlScanControl::Continue)
    })?;

    flush_validation_ranges(
        validation_ranges,
        validations,
        store,
        &mut worksheet,
        Some("content.xml"),
    );

    Ok(worksheet)
}

fn handle_table_start_full(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    ctx: FullTableContext<'_>,
) -> Result<(), ParseError> {
    let FullTableContext {
        store,
        limits,
        style_map,
        next_style_id,
        row_idx,
        worksheet,
        validation_ranges,
        cell_values,
        formula_cells,
        formula_map,
        shapes,
    } = ctx;
    match local_name(start.name().as_ref()) {
        b"table-row" => {
            let row_repeat = row_repeat_from(start);
            limits.bump_rows(row_repeat as u64)?;
            let row_cells = parse_ods_row(reader, start, store, style_map, next_style_id)?;
            for _ in 0..row_repeat {
                let mut row_state = RowBuildState {
                    validation_ranges,
                    cell_values,
                    formula_cells,
                    formula_map,
                };
                emit_full_row_cells(
                    &row_cells,
                    *row_idx,
                    store,
                    worksheet,
                    limits,
                    &mut row_state,
                )?;
                *row_idx += 1;
            }
        }
        b"frame" => {
            if let Some(shape_id) = spreadsheet::parse_draw_frame_spreadsheet(reader, start, store)?
            {
                shapes.push(shape_id);
            }
        }
        b"conditional-formatting" => {
            if let Some(cf) = parse_ods_conditional_formatting(reader, start)? {
                push_conditional_format(store, worksheet, cf);
            }
        }
        b"filter" | b"filter-and" | b"filter-or" => {
            spreadsheet::skip_element(reader, start.name().as_ref())?;
        }
        _ => {}
    }
    Ok(())
}

fn handle_table_empty_full(
    reader: &mut OdfReader<'_>,
    empty: &BytesStart<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    row_idx: &mut u32,
    worksheet: &mut Worksheet,
    shapes: &mut Vec<NodeId>,
) -> Result<(), ParseError> {
    match local_name(empty.name().as_ref()) {
        b"table-row" => {
            let row_repeat = row_repeat_from(empty);
            limits.bump_rows(row_repeat as u64)?;
            *row_idx += row_repeat;
        }
        b"frame" => {
            if let Some(shape_id) = spreadsheet::parse_draw_frame_spreadsheet(reader, empty, store)?
            {
                shapes.push(shape_id);
            }
        }
        b"conditional-formatting" => {
            if let Some(cf) = parse_ods_conditional_formatting_empty(empty)? {
                push_conditional_format(store, worksheet, cf);
            }
        }
        _ => {}
    }
    Ok(())
}

fn handle_table_start_fast(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    sample_rows: u32,
    sample_cols: u32,
    sample_enabled: bool,
    ctx: FastTableContext<'_>,
) -> Result<(), ParseError> {
    let FastTableContext {
        store,
        limits,
        style_map,
        next_style_id,
        row_idx,
        worksheet,
        validation_ranges,
    } = ctx;
    match local_name(start.name().as_ref()) {
        b"table-row" => {
            let row_repeat = row_repeat_from(start);
            limits.bump_rows(row_repeat as u64)?;
            if sample_enabled && *row_idx < sample_rows {
                let row_cells = parse_ods_row_sample(
                    reader,
                    start,
                    store,
                    style_map,
                    next_style_id,
                    sample_cols,
                )?;
                let remaining = sample_rows.saturating_sub(*row_idx);
                let repeat = row_repeat.min(remaining);
                for _ in 0..repeat {
                    emit_sampled_row_cells(
                        &row_cells,
                        *row_idx,
                        sample_cols,
                        store,
                        worksheet,
                        limits,
                        validation_ranges,
                    )?;
                    *row_idx += 1;
                }
                if row_repeat > repeat {
                    *row_idx += row_repeat - repeat;
                }
            } else {
                spreadsheet::skip_element(reader, start.name().as_ref())?;
                *row_idx += row_repeat;
            }
        }
        b"frame" => {
            spreadsheet::skip_element(reader, start.name().as_ref())?;
        }
        _ => {}
    }
    Ok(())
}

fn handle_table_empty_fast(
    empty: &BytesStart<'_>,
    limits: &dyn OdfLimitCounter,
    row_idx: &mut u32,
) -> Result<(), ParseError> {
    if local_name(empty.name().as_ref()) == b"table-row" {
        let row_repeat = row_repeat_from(empty);
        limits.bump_rows(row_repeat as u64)?;
        *row_idx += row_repeat;
    }
    Ok(())
}

pub(crate) fn parse_ods_cell(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    _store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<OdsCellData, ParseError> {
    let mut value_type = attr_value_by_suffix(start, &[b":cell-value-type"])
        .or_else(|| attr_value_by_suffix(start, &[b":value-type"]));
    let mut value_attr = attr_value_by_suffix(start, &[b":cell-value"])
        .or_else(|| attr_value_by_suffix(start, &[b":value"]));
    let date_value = attr_value_by_suffix(start, &[b":date-value"]);
    let time_value = attr_value_by_suffix(start, &[b":time-value"]);
    let formula_attr = attr_value_by_suffix(start, &[b":formula"]);
    let col_repeat = attr_value_by_suffix(start, &[b":number-columns-repeated"])
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let col_span = attr_value_by_suffix(start, &[b":number-columns-spanned"])
        .and_then(|v| v.parse::<u32>().ok());
    let row_span =
        attr_value_by_suffix(start, &[b":number-rows-spanned"]).and_then(|v| v.parse::<u32>().ok());
    let validation_name = attr_value_by_suffix(start, &[b":content-validation-name"]);

    let style_id = resolve_style_id(start, style_map, next_style_id);
    let text = read_ods_cell_text(reader)?;
    infer_cell_value_type_and_attr(
        &mut value_type,
        &mut value_attr,
        date_value,
        time_value,
        &text,
    );

    let value = parse_cell_value_with_text(value_type.as_deref(), value_attr.as_deref(), &text);
    let formula = parse_cell_formula(formula_attr);

    Ok(OdsCellData {
        value,
        formula,
        style_id,
        col_repeat,
        validation_name,
        col_span,
        row_span,
        is_covered: false,
    })
}

pub(crate) fn parse_ods_cell_empty(
    start: &BytesStart<'_>,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<OdsCellData, ParseError> {
    let value_type = attr_value_by_suffix(start, &[b":cell-value-type"])
        .or_else(|| attr_value_by_suffix(start, &[b":value-type"]));
    let value_attr = attr_value_by_suffix(start, &[b":cell-value"])
        .or_else(|| attr_value_by_suffix(start, &[b":value"]))
        .or_else(|| attr_value_by_suffix(start, &[b":date-value"]))
        .or_else(|| attr_value_by_suffix(start, &[b":time-value"]));
    let formula_attr = attr_value_by_suffix(start, &[b":formula"]);
    let col_repeat = attr_value_by_suffix(start, &[b":number-columns-repeated"])
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let col_span = attr_value_by_suffix(start, &[b":number-columns-spanned"])
        .and_then(|v| v.parse::<u32>().ok());
    let row_span =
        attr_value_by_suffix(start, &[b":number-rows-spanned"]).and_then(|v| v.parse::<u32>().ok());
    let validation_name = attr_value_by_suffix(start, &[b":content-validation-name"]);
    let style_id = resolve_style_id(start, style_map, next_style_id);
    let value = parse_cell_value_empty(value_type.as_deref(), value_attr.as_deref());
    let formula = parse_cell_formula(formula_attr);

    Ok(OdsCellData {
        value,
        formula,
        style_id,
        col_repeat,
        validation_name,
        col_span,
        row_span,
        is_covered: false,
    })
}
