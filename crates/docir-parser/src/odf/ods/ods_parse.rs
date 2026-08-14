use super::super::helpers::{
    OdsCellData, ValidationDef, parse_ods_conditional_formatting,
    parse_ods_conditional_formatting_empty, parse_ods_row,
};
use super::ods_postprocess::emit_full_row_cells;
use super::{
    RowBuildState, attach_shapes_as_drawing, emit_sampled_row_cells, flush_validation_ranges,
    infer_cell_value_type_and_attr, parse_cell_formula, parse_cell_value_empty,
    parse_cell_value_with_text, push_conditional_format, read_ods_cell_text, resolve_style_id,
    row_repeat_from,
};
use crate::odf::{
    CellValue, IrStore, NodeId, OdfLimitCounter, OdfReader, ParseError, Worksheet,
    evaluate_ods_formulas, parse_ods_row_sample, spreadsheet,
};
use crate::xml_utils::{
    XmlScanControl, dispatch_start_or_empty, is_end_event_local, local_name,
    scan_xml_events_with_reader, try_attr_value_by_suffix, xml_error,
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

fn advance_row_index(row_idx: &mut u32, rows: u32) -> Result<(), ParseError> {
    *row_idx = row_idx
        .checked_add(rows)
        .ok_or_else(|| ParseError::InvalidStructure("ODF row index overflow".to_string()))?;
    Ok(())
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
    let name = try_attr_value_by_suffix(start, &[b":name"], "content.xml")?
        .unwrap_or_else(|| format!("Sheet{sheet_id}"));
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
    let mut nested_table_depth = 0usize;
    let mut reached_table_end = false;

    scan_xml_events_with_reader(reader, &mut buf, "content.xml", |reader, event| {
        if is_end_event_local(&event, b"table") {
            if nested_table_depth == 0 {
                reached_table_end = true;
                return Ok(XmlScanControl::Break);
            }
            nested_table_depth -= 1;
        } else if matches!(&event, quick_xml::events::Event::Start(start) if local_name(start.name().as_ref()) == b"table")
        {
            nested_table_depth = nested_table_depth
                .checked_add(1)
                .ok_or_else(|| xml_error("content.xml", "ODF table nesting depth overflow"))?;
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
                handle_table_empty_full(e, store, limits, &mut row_idx, &mut worksheet)
            }
        })?;
        Ok(XmlScanControl::Continue)
    })?;
    if !reached_table_end {
        return Err(xml_error(
            "content.xml",
            "unexpected end of XML before closing table",
        ));
    }

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
    let name = try_attr_value_by_suffix(start, &[b":name"], "content.xml")?
        .unwrap_or_else(|| format!("Sheet{sheet_id}"));
    let mut worksheet = Worksheet::new(name, sheet_id);
    let mut buf = Vec::new();
    let mut row_idx: u32 = 0;
    let mut style_map: HashMap<String, u32> = HashMap::new();
    let mut next_style_id = 1u32;
    let mut validation_ranges: HashMap<String, Vec<String>> = HashMap::new();
    let sample_rows = limits.sample_rows();
    let sample_cols = limits.sample_cols();
    let sample_enabled = sample_rows > 0 && sample_cols > 0;
    let mut nested_table_depth = 0usize;
    let mut reached_table_end = false;

    scan_xml_events_with_reader(reader, &mut buf, "content.xml", |reader, event| {
        if is_end_event_local(&event, b"table") {
            if nested_table_depth == 0 {
                reached_table_end = true;
                return Ok(XmlScanControl::Break);
            }
            nested_table_depth -= 1;
        } else if matches!(&event, quick_xml::events::Event::Start(start) if local_name(start.name().as_ref()) == b"table")
        {
            nested_table_depth = nested_table_depth
                .checked_add(1)
                .ok_or_else(|| xml_error("content.xml", "ODF table nesting depth overflow"))?;
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
    if !reached_table_end {
        return Err(xml_error(
            "content.xml",
            "unexpected end of XML before closing table",
        ));
    }

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
            let row_repeat = row_repeat_from(start)?;
            limits.bump_rows(row_repeat as u64)?;
            let row_cells = parse_ods_row(reader, start, store, style_map, next_style_id)?;
            if row_cells.cells.is_empty() {
                advance_row_index(row_idx, row_repeat)?;
                return Ok(());
            }
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
                advance_row_index(row_idx, 1)?;
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
    empty: &BytesStart<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    row_idx: &mut u32,
    worksheet: &mut Worksheet,
) -> Result<(), ParseError> {
    match local_name(empty.name().as_ref()) {
        b"table-row" => {
            let row_repeat = row_repeat_from(empty)?;
            limits.bump_rows(row_repeat as u64)?;
            advance_row_index(row_idx, row_repeat)?;
        }
        b"frame" => {}
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
            let row_repeat = row_repeat_from(start)?;
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
                    advance_row_index(row_idx, 1)?;
                }
                if row_repeat > repeat {
                    advance_row_index(row_idx, row_repeat - repeat)?;
                }
            } else {
                spreadsheet::skip_element(reader, start.name().as_ref())?;
                advance_row_index(row_idx, row_repeat)?;
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
        let row_repeat = row_repeat_from(empty)?;
        limits.bump_rows(row_repeat as u64)?;
        advance_row_index(row_idx, row_repeat)?;
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
    let mut value_type = ods_cell_attr(start, &[b":cell-value-type", b":value-type"])?;
    let mut value_attr = ods_cell_attr(start, &[b":cell-value", b":value"])?;
    let date_value = ods_cell_attr(start, &[b":date-value"])?;
    let time_value = ods_cell_attr(start, &[b":time-value"])?;
    let formula_attr = ods_cell_attr(start, &[b":formula"])?;
    let col_repeat = ods_cell_u32_attr(start, b":number-columns-repeated")?.unwrap_or(1);
    let col_span = ods_cell_u32_attr(start, b":number-columns-spanned")?;
    let row_span = ods_cell_u32_attr(start, b":number-rows-spanned")?;
    let validation_name = ods_cell_attr(start, &[b":content-validation-name"])?;

    let style_id = resolve_style_id(start, style_map, next_style_id)?;
    let text = read_ods_cell_text(reader)?;
    infer_cell_value_type_and_attr(
        &mut value_type,
        &mut value_attr,
        date_value,
        time_value,
        &text,
    );

    let value = parse_cell_value_with_text(value_type.as_deref(), value_attr.as_deref(), &text)?;
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
    let value_type = ods_cell_attr(start, &[b":cell-value-type", b":value-type"])?;
    let value_attr = ods_cell_attr(
        start,
        &[b":cell-value", b":value", b":date-value", b":time-value"],
    )?;
    let formula_attr = ods_cell_attr(start, &[b":formula"])?;
    let col_repeat = ods_cell_u32_attr(start, b":number-columns-repeated")?.unwrap_or(1);
    let col_span = ods_cell_u32_attr(start, b":number-columns-spanned")?;
    let row_span = ods_cell_u32_attr(start, b":number-rows-spanned")?;
    let validation_name = ods_cell_attr(start, &[b":content-validation-name"])?;
    let style_id = resolve_style_id(start, style_map, next_style_id)?;
    let value = parse_cell_value_empty(value_type.as_deref(), value_attr.as_deref())?;
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

fn ods_cell_attr(start: &BytesStart<'_>, suffixes: &[&[u8]]) -> Result<Option<String>, ParseError> {
    for suffix in suffixes {
        if let Some(value) = try_attr_value_by_suffix(start, &[*suffix], "content.xml")? {
            return Ok(Some(value));
        }
    }
    Ok(None)
}

fn ods_cell_u32_attr(start: &BytesStart<'_>, suffix: &[u8]) -> Result<Option<u32>, ParseError> {
    ods_cell_attr(start, &[suffix])?
        .map(|v| {
            let parsed = v.parse().map_err(|err| xml_error("content.xml", err))?;
            if parsed == 0 {
                return Err(xml_error(
                    "content.xml",
                    "ODF cell span/repeat attributes must be positive",
                ));
            }
            Ok(parsed)
        })
        .transpose()
}

#[cfg(test)]
mod tests {
    use super::advance_row_index;
    use crate::error::ParseError;

    #[test]
    fn advance_row_index_rejects_overflow() {
        let mut row_idx = u32::MAX;

        let err = advance_row_index(&mut row_idx, 1).expect_err("row index must not wrap");

        assert!(
            matches!(err, ParseError::InvalidStructure(message) if message.contains("row index overflow"))
        );
        assert_eq!(row_idx, u32::MAX);
    }
}
