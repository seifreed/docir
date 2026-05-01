use super::super::helpers::{column_index_to_name, OdsRow, ValidationDef};
use super::RowBuildState;
use crate::odf::{
    Cell, ConditionalFormat, DataValidation, IRNode, IrStore, NodeId, OdfLimitCounter, ParseError,
    SourceSpan, Worksheet, WorksheetDrawing,
};
use std::collections::HashMap;

pub(crate) fn push_conditional_format(
    store: &mut IrStore,
    worksheet: &mut Worksheet,
    cf: ConditionalFormat,
) {
    let id = cf.id;
    store.insert(IRNode::ConditionalFormat(cf));
    worksheet.conditional_formats.push(id);
}

pub(crate) fn flush_validation_ranges(
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

pub(crate) fn attach_shapes_as_drawing(
    shapes: Vec<NodeId>,
    store: &mut IrStore,
    worksheet: &mut Worksheet,
) {
    if shapes.is_empty() {
        return;
    }
    let mut drawing = WorksheetDrawing::new();
    drawing.shapes = shapes;
    let drawing_id = drawing.id;
    store.insert(IRNode::WorksheetDrawing(drawing));
    worksheet.drawings.push(drawing_id);
}

pub(super) fn emit_full_row_cells(
    row_cells: &OdsRow,
    row_idx: u32,
    store: &mut IrStore,
    worksheet: &mut Worksheet,
    limits: &dyn OdfLimitCounter,
    row_state: &mut RowBuildState<'_>,
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
                row_state
                    .cell_values
                    .insert((row_idx, col_idx), cell_data.value.clone());
                if let Some(formula) = &cell_data.formula {
                    row_state
                        .formula_map
                        .insert((row_idx, col_idx), formula.text.clone());
                    row_state
                        .formula_cells
                        .push((cell_id, row_idx, col_idx, formula.text.clone()));
                }
                if let Some(validation) = &cell_data.validation_name {
                    row_state
                        .validation_ranges
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

pub(crate) fn emit_sampled_row_cells(
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
