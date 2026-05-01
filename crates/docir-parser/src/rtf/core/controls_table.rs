use super::super::core_parse::set_border_target;
use super::super::state::{BorderTarget, RtfParseContext};
use super::super::{
    apply_border, apply_paragraph_border, color_from_index, ensure_section, finalize_cell,
    finalize_row, flush_text,
};
use crate::error::ParseError;
use docir_core::ir::{
    BorderStyle, CellVerticalAlignment, MergeType, Table, TableCell, TableCellProperties, TableRow,
};
use docir_core::visitor::IrStore;

fn set_pending_cell_vertical_align(ctx: &mut RtfParseContext, value: CellVerticalAlignment) {
    ctx.pending_cell_props
        .get_or_insert_with(TableCellProperties::default)
        .vertical_align = Some(value);
}

fn set_pending_border_style(ctx: &mut RtfParseContext, style: BorderStyle) {
    ctx.pending_border.style = style;
    apply_border(ctx);
}

pub(super) fn handle_table_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<bool, ParseError> {
    if handle_table_border_controls(word, param, ctx) {
        return Ok(true);
    }
    if handle_table_cell_property_controls(word, param, ctx) {
        return Ok(true);
    }
    match word {
        "trowd" => {
            flush_text(ctx, store, None)?;
            ensure_section(ctx, store);
            if ctx.current_table.is_none() {
                ctx.current_table = Some(Table::new());
            }
            ctx.current_row = Some(TableRow::new());
            ctx.current_cell = Some(TableCell::new());
            ctx.row_cellx.clear();
            ctx.current_cell_index = 0;
            ctx.pending_cell_props = None;
            ctx.pending_border_target = None;
        }
        "cellx" => {
            if let Some(value) = param {
                ctx.row_cellx.push(value);
            }
        }
        "cell" => {
            flush_text(ctx, store, None)?;
            finalize_cell(ctx, store);
            ctx.current_cell = Some(TableCell::new());
        }
        "row" => {
            flush_text(ctx, store, None)?;
            finalize_cell(ctx, store);
            finalize_row(ctx, store);
        }
        _ => return Ok(false),
    }
    Ok(true)
}

pub(super) fn handle_table_cell_property_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
) -> bool {
    match word {
        "clvmgf" => {
            ctx.pending_cell_props
                .get_or_insert_with(TableCellProperties::default)
                .vertical_merge = Some(MergeType::Restart);
        }
        "clvmrg" => {
            ctx.pending_cell_props
                .get_or_insert_with(TableCellProperties::default)
                .vertical_merge = Some(MergeType::Continue);
        }
        "clgridspan" => {
            if let Some(value) = param {
                ctx.pending_cell_props
                    .get_or_insert_with(TableCellProperties::default)
                    .grid_span = Some(value.max(1) as u32);
            }
        }
        "clcbpat" => {
            if let Some(index) = param {
                if let Some(color) = color_from_index(ctx, index as usize) {
                    ctx.pending_cell_props
                        .get_or_insert_with(TableCellProperties::default)
                        .shading = Some(color);
                }
            }
        }
        "clvertalt" => set_pending_cell_vertical_align(ctx, CellVerticalAlignment::Top),
        "clvertalc" => set_pending_cell_vertical_align(ctx, CellVerticalAlignment::Center),
        "clvertalb" => set_pending_cell_vertical_align(ctx, CellVerticalAlignment::Bottom),
        _ => return false,
    }
    true
}

pub(super) fn handle_table_border_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
) -> bool {
    match word {
        "clbrdrt" => set_border_target(ctx, BorderTarget::Top),
        "clbrdrb" => set_border_target(ctx, BorderTarget::Bottom),
        "clbrdrl" => set_border_target(ctx, BorderTarget::Left),
        "clbrdrr" => set_border_target(ctx, BorderTarget::Right),
        "clbrdrh" => set_border_target(ctx, BorderTarget::InsideH),
        "clbrdrv" => set_border_target(ctx, BorderTarget::InsideV),
        "brdrs" => set_pending_border_style(ctx, BorderStyle::Single),
        "brdrth" => set_pending_border_style(ctx, BorderStyle::Thick),
        "brdrdb" => set_pending_border_style(ctx, BorderStyle::Double),
        "brdrdot" => set_pending_border_style(ctx, BorderStyle::Dotted),
        "brdrdash" => set_pending_border_style(ctx, BorderStyle::Dashed),
        "brdrtriple" => set_pending_border_style(ctx, BorderStyle::Triple),
        "brdrw" => {
            if let Some(value) = param {
                ctx.pending_border.width = Some(value.max(0) as u32);
                apply_border(ctx);
            }
        }
        "brdrcf" => {
            if let Some(index) = param {
                if let Some(color) = color_from_index(ctx, index as usize) {
                    ctx.pending_border.color = Some(color);
                    apply_border(ctx);
                }
            }
        }
        "brdrt" => {
            ctx.pending_para_border_target = Some(BorderTarget::Top);
            apply_paragraph_border(ctx);
        }
        "brdrb" => {
            ctx.pending_para_border_target = Some(BorderTarget::Bottom);
            apply_paragraph_border(ctx);
        }
        "brdrl" => {
            ctx.pending_para_border_target = Some(BorderTarget::Left);
            apply_paragraph_border(ctx);
        }
        "brdrr" => {
            ctx.pending_para_border_target = Some(BorderTarget::Right);
            apply_paragraph_border(ctx);
        }
        _ => return false,
    }
    true
}
