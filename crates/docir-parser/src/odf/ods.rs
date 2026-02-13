use super::*;

pub(super) fn parse_ods_table(
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
    let name = attr_value(start, b"table:name").unwrap_or_else(|| format!("Sheet{sheet_id}"));
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

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => handle_table_start_full(
                reader,
                &e,
                store,
                limits,
                &mut style_map,
                &mut next_style_id,
                &mut row_idx,
                &mut worksheet,
                &mut validation_ranges,
                &mut cell_values,
                &mut formula_cells,
                &mut formula_map,
                &mut shapes,
            )?,
            Ok(Event::Empty(e)) => handle_table_empty_full(
                reader,
                &e,
                store,
                limits,
                &mut row_idx,
                &mut worksheet,
                &mut shapes,
            )?,
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(crate::xml_utils::xml_error("content.xml", e));
            }
            _ => {}
        }
        buf.clear();
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

pub(super) fn parse_ods_table_fast(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    sheet_id: u32,
    store: &mut IrStore,
    validations: &HashMap<String, ValidationDef>,
    limits: &dyn OdfLimitCounter,
) -> Result<Worksheet, ParseError> {
    let name = attr_value(start, b"table:name").unwrap_or_else(|| format!("Sheet{sheet_id}"));
    let mut worksheet = Worksheet::new(name, sheet_id);
    let mut buf = Vec::new();
    let mut row_idx: u32 = 0;
    let mut style_map: HashMap<String, u32> = HashMap::new();
    let mut next_style_id = 1u32;
    let mut validation_ranges: HashMap<String, Vec<String>> = HashMap::new();
    let sample_rows = limits.sample_rows();
    let sample_cols = limits.sample_cols();
    let sample_enabled = sample_rows > 0 && sample_cols > 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => handle_table_start_fast(
                reader,
                &e,
                store,
                limits,
                sample_rows,
                sample_cols,
                sample_enabled,
                &mut style_map,
                &mut next_style_id,
                &mut row_idx,
                &mut worksheet,
                &mut validation_ranges,
            )?,
            Ok(Event::Empty(e)) => handle_table_empty_fast(&e, limits, &mut row_idx)?,
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(crate::xml_utils::xml_error("content.xml", e));
            }
            _ => {}
        }
        buf.clear();
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

#[allow(clippy::too_many_arguments)]
fn handle_table_start_full(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
    row_idx: &mut u32,
    worksheet: &mut Worksheet,
    validation_ranges: &mut HashMap<String, Vec<String>>,
    cell_values: &mut HashMap<(u32, u32), CellValue>,
    formula_cells: &mut Vec<(NodeId, u32, u32, String)>,
    formula_map: &mut HashMap<(u32, u32), String>,
    shapes: &mut Vec<NodeId>,
) -> Result<(), ParseError> {
    match start.name().as_ref() {
        b"table:table-row" => {
            let row_repeat = row_repeat_from(start);
            limits.bump_rows(row_repeat as u64)?;
            let row_cells = parse_ods_row(reader, start, store, style_map, next_style_id)?;
            for _ in 0..row_repeat {
                emit_full_row_cells(
                    &row_cells,
                    *row_idx,
                    store,
                    worksheet,
                    limits,
                    validation_ranges,
                    cell_values,
                    formula_cells,
                    formula_map,
                )?;
                *row_idx += 1;
            }
        }
        b"draw:frame" => {
            if let Some(shape_id) = spreadsheet::parse_draw_frame_spreadsheet(reader, start, store)?
            {
                shapes.push(shape_id);
            }
        }
        b"table:conditional-formatting" => {
            if let Some(cf) = parse_ods_conditional_formatting(reader, start)? {
                push_conditional_format(store, worksheet, cf);
            }
        }
        b"table:filter" | b"table:filter-and" | b"table:filter-or" => {
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
    match empty.name().as_ref() {
        b"table:table-row" => {
            let row_repeat = row_repeat_from(empty);
            limits.bump_rows(row_repeat as u64)?;
            *row_idx += row_repeat;
        }
        b"draw:frame" => {
            if let Some(shape_id) = spreadsheet::parse_draw_frame_spreadsheet(reader, empty, store)?
            {
                shapes.push(shape_id);
            }
        }
        b"table:conditional-formatting" => {
            if let Some(cf) = parse_ods_conditional_formatting_empty(empty)? {
                push_conditional_format(store, worksheet, cf);
            }
        }
        _ => {}
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn handle_table_start_fast(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    sample_rows: u32,
    sample_cols: u32,
    sample_enabled: bool,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
    row_idx: &mut u32,
    worksheet: &mut Worksheet,
    validation_ranges: &mut HashMap<String, Vec<String>>,
) -> Result<(), ParseError> {
    match start.name().as_ref() {
        b"table:table-row" => {
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
        b"draw:frame" => {
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
    if empty.name().as_ref() == b"table:table-row" {
        let row_repeat = row_repeat_from(empty);
        limits.bump_rows(row_repeat as u64)?;
        *row_idx += row_repeat;
    }
    Ok(())
}

fn row_repeat_from(start: &BytesStart<'_>) -> u32 {
    attr_value(start, b"table:number-rows-repeated")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1)
}

fn push_conditional_format(store: &mut IrStore, worksheet: &mut Worksheet, cf: ConditionalFormat) {
    let id = cf.id;
    store.insert(IRNode::ConditionalFormat(cf));
    worksheet.conditional_formats.push(id);
}

fn flush_validation_ranges(
    validation_ranges: HashMap<String, Vec<String>>,
    validations: &HashMap<String, ValidationDef>,
    store: &mut IrStore,
    worksheet: &mut Worksheet,
    span_path: Option<&str>,
) {
    for (name, ranges) in validation_ranges {
        let mut validation = build_data_validation(validations.get(&name), ranges);
        validation.span = span_path.map(SourceSpan::new);
        let validation_id = validation.id;
        store.insert(IRNode::DataValidation(validation));
        worksheet.data_validations.push(validation_id);
    }
}

fn build_data_validation(def: Option<&ValidationDef>, ranges: Vec<String>) -> DataValidation {
    match def {
        Some(def) => DataValidation {
            id: NodeId::new(),
            validation_type: def.validation_type.clone(),
            operator: def.operator.clone(),
            allow_blank: def.allow_blank,
            show_input_message: def.show_input_message,
            show_error_message: def.show_error_message,
            error_title: def.error_title.clone(),
            error: def.error.clone(),
            prompt_title: def.prompt_title.clone(),
            prompt: def.prompt.clone(),
            ranges,
            formula1: def.formula1.clone(),
            formula2: def.formula2.clone(),
            span: None,
        },
        None => DataValidation {
            id: NodeId::new(),
            validation_type: None,
            operator: None,
            allow_blank: false,
            show_input_message: false,
            show_error_message: false,
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
            ranges,
            formula1: None,
            formula2: None,
            span: None,
        },
    }
}

fn attach_shapes_as_drawing(shapes: Vec<NodeId>, store: &mut IrStore, worksheet: &mut Worksheet) {
    if shapes.is_empty() {
        return;
    }
    let mut drawing = WorksheetDrawing::new();
    drawing.shapes = shapes;
    let drawing_id = drawing.id;
    store.insert(IRNode::WorksheetDrawing(drawing));
    worksheet.drawings.push(drawing_id);
}

fn emit_full_row_cells(
    row_cells: &OdsRow,
    row_idx: u32,
    store: &mut IrStore,
    worksheet: &mut Worksheet,
    limits: &dyn OdfLimitCounter,
    validation_ranges: &mut HashMap<String, Vec<String>>,
    cell_values: &mut HashMap<(u32, u32), CellValue>,
    formula_cells: &mut Vec<(NodeId, u32, u32, String)>,
    formula_map: &mut HashMap<(u32, u32), String>,
) -> Result<(), ParseError> {
    let mut col_idx: u32 = 0;
    for cell_data in &row_cells.cells {
        for _ in 0..cell_data.col_repeat {
            if cell_data.is_covered {
                col_idx += 1;
                continue;
            }
            if let Some(range) = cell_data.merge_range(row_idx, col_idx) {
                worksheet.merged_cells.push(range);
            }
            if cell_data.should_emit() {
                limits.bump_cells(1)?;
                let reference = format!("{}{}", column_index_to_name(col_idx), row_idx + 1);
                let mut cell = Cell::new(reference.clone(), col_idx, row_idx);
                cell.value = cell_data.value.clone();
                cell.formula = cell_data.formula.clone();
                cell.style_id = cell_data.style_id;
                let cell_id = cell.id;
                store.insert(IRNode::Cell(cell));
                worksheet.cells.push(cell_id);
                cell_values.insert((row_idx, col_idx), cell_data.value.clone());
                if let Some(formula) = &cell_data.formula {
                    formula_map.insert((row_idx, col_idx), formula.text.clone());
                    formula_cells.push((cell_id, row_idx, col_idx, formula.text.clone()));
                }
                if let Some(validation) = &cell_data.validation_name {
                    validation_ranges
                        .entry(validation.clone())
                        .or_default()
                        .push(reference);
                }
            }
            col_idx += 1;
        }
    }
    Ok(())
}

fn emit_sampled_row_cells(
    row_cells: &OdsRow,
    row_idx: u32,
    sample_cols: u32,
    store: &mut IrStore,
    worksheet: &mut Worksheet,
    limits: &dyn OdfLimitCounter,
    validation_ranges: &mut HashMap<String, Vec<String>>,
) -> Result<(), ParseError> {
    let mut col_idx: u32 = 0;
    for cell_data in &row_cells.cells {
        for _ in 0..cell_data.col_repeat {
            if col_idx >= sample_cols {
                break;
            }
            if cell_data.is_covered {
                col_idx += 1;
                continue;
            }
            if let Some(range) = cell_data.merge_range(row_idx, col_idx) {
                worksheet.merged_cells.push(range);
            }
            if cell_data.should_emit() {
                limits.bump_cells(1)?;
                let reference = format!("{}{}", column_index_to_name(col_idx), row_idx + 1);
                let mut cell = Cell::new(reference.clone(), col_idx, row_idx);
                cell.value = cell_data.value.clone();
                cell.formula = cell_data.formula.clone();
                cell.style_id = cell_data.style_id;
                let cell_id = cell.id;
                store.insert(IRNode::Cell(cell));
                worksheet.cells.push(cell_id);
                if let Some(validation) = &cell_data.validation_name {
                    validation_ranges
                        .entry(validation.clone())
                        .or_default()
                        .push(reference);
                }
            }
            col_idx += 1;
        }
        if col_idx >= sample_cols {
            break;
        }
    }
    Ok(())
}

pub(super) fn parse_ods_cell(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    _store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<OdsCellData, ParseError> {
    let mut value_type = attr_value(start, b"table:cell-value-type")
        .or_else(|| attr_value(start, b"office:value-type"));
    let mut value_attr =
        attr_value(start, b"table:cell-value").or_else(|| attr_value(start, b"office:value"));
    let date_value = attr_value(start, b"table:date-value");
    let time_value = attr_value(start, b"table:time-value");
    let formula_attr = attr_value(start, b"table:formula");
    let col_repeat = attr_value(start, b"table:number-columns-repeated")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let col_span =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok());
    let row_span =
        attr_value(start, b"table:number-rows-spanned").and_then(|v| v.parse::<u32>().ok());
    let validation_name = attr_value(start, b"table:content-validation-name");

    let style_id = resolve_style_id(start, style_map, next_style_id);

    let mut text = String::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:p" => {
                    let para = parse_text_element(reader, e.name().as_ref())?;
                    if !text.is_empty() && !para.is_empty() {
                        text.push('\n');
                    }
                    text.push_str(&para);
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                let chunk = e.unescape().unwrap_or_default();
                text.push_str(&chunk);
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table-cell" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(crate::xml_utils::xml_error("content.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    if value_type.is_none() && !text.is_empty() {
        value_type = Some("string".to_string());
    }

    if value_attr.is_none() {
        value_attr = date_value.or(time_value).or_else(|| {
            if !text.is_empty() {
                Some(text.clone())
            } else {
                None
            }
        });
    }

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

pub(super) fn parse_ods_cell_empty(
    start: &BytesStart<'_>,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<OdsCellData, ParseError> {
    let value_type = attr_value(start, b"table:cell-value-type")
        .or_else(|| attr_value(start, b"office:value-type"));
    let value_attr = attr_value(start, b"table:cell-value")
        .or_else(|| attr_value(start, b"office:value"))
        .or_else(|| attr_value(start, b"table:date-value"))
        .or_else(|| attr_value(start, b"table:time-value"));
    let formula_attr = attr_value(start, b"table:formula");
    let col_repeat = attr_value(start, b"table:number-columns-repeated")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let col_span =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok());
    let row_span =
        attr_value(start, b"table:number-rows-spanned").and_then(|v| v.parse::<u32>().ok());
    let validation_name = attr_value(start, b"table:content-validation-name");
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

fn resolve_style_id(
    start: &BytesStart<'_>,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Option<u32> {
    attr_value(start, b"table:style-name").map(|name| {
        if let Some(id) = style_map.get(&name) {
            *id
        } else {
            let id = *next_style_id;
            *next_style_id += 1;
            style_map.insert(name, id);
            id
        }
    })
}

fn parse_cell_formula(formula_attr: Option<String>) -> Option<CellFormula> {
    formula_attr.map(|formula| CellFormula {
        text: normalize_formula_text(formula),
        formula_type: docir_core::ir::FormulaType::Normal,
        shared_index: None,
        shared_ref: None,
        is_array: false,
        array_ref: None,
    })
}

fn normalize_formula_text(formula: String) -> String {
    if let Some(stripped) = formula.strip_prefix("of:=") {
        stripped.to_string()
    } else if let Some(stripped) = formula.strip_prefix("of:") {
        stripped.to_string()
    } else if let Some(stripped) = formula.strip_prefix('=') {
        stripped.to_string()
    } else {
        formula
    }
}

fn parse_cell_value_with_text(
    value_type: Option<&str>,
    value_attr: Option<&str>,
    text: &str,
) -> CellValue {
    match value_type {
        Some("string") | Some("text") => {
            if !text.is_empty() {
                CellValue::String(text.to_string())
            } else {
                CellValue::String(value_attr.unwrap_or_default().to_string())
            }
        }
        Some("float") | Some("currency") | Some("percentage") => value_attr
            .and_then(|v| v.parse::<f64>().ok())
            .map(CellValue::Number)
            .unwrap_or_else(|| CellValue::String(text.to_string())),
        Some("boolean") => {
            let v = value_attr.unwrap_or(text);
            CellValue::Boolean(v == "true" || v == "1")
        }
        Some("date") | Some("time") => CellValue::String(value_attr.unwrap_or(text).to_string()),
        _ => {
            if !text.is_empty() {
                CellValue::String(text.to_string())
            } else {
                CellValue::Empty
            }
        }
    }
}

fn parse_cell_value_empty(value_type: Option<&str>, value_attr: Option<&str>) -> CellValue {
    match value_type {
        Some("float") | Some("currency") | Some("percentage") => value_attr
            .and_then(|v| v.parse::<f64>().ok())
            .map(CellValue::Number)
            .unwrap_or(CellValue::Empty),
        Some("boolean") => {
            let v = value_attr.unwrap_or_default();
            if v.is_empty() {
                CellValue::Empty
            } else {
                CellValue::Boolean(v == "true" || v == "1")
            }
        }
        Some("string") | Some("text") | Some("date") | Some("time") => value_attr
            .map_or(CellValue::Empty, |value| {
                CellValue::String(value.to_string())
            }),
        _ => CellValue::Empty,
    }
}
