use super::state::{BorderTarget, GroupKind, RtfParseContext, StyleEntryContext};
use super::{
    apply_border, apply_paragraph_border, color_from_index, ensure_section, finalize_cell,
    finalize_row, flush_text, pending_numbering, set_border_target, BorderStyle,
    CellVerticalAlignment, MergeType, ObjectContext, ObjectTextTarget, StyleSet, StyleType, Table,
    TableCell, TableCellProperties, TableRow,
};
use crate::error::ParseError;
use docir_core::visitor::IrStore;

fn set_last_group_kind(ctx: &mut RtfParseContext, kind: GroupKind) {
    if let Some(group) = ctx.group_stack.last_mut() {
        group.kind = kind;
    }
}

fn set_stylesheet_entry(ctx: &mut RtfParseContext, style_id: String, style_type: StyleType) {
    set_last_group_kind(ctx, GroupKind::StylesheetEntry);
    ctx.current_style = Some(StyleEntryContext {
        style_id: Some(style_id),
        style_type: Some(style_type),
        name_buf: String::new(),
    });
}

fn apply_pending_numbering_to_current_paragraph(ctx: &mut RtfParseContext) {
    let numbering = pending_numbering(ctx);
    if let Some(para) = ctx.current_paragraph.as_mut() {
        if let Some(numbering) = numbering {
            para.properties.numbering = Some(numbering);
        }
    }
}

fn set_last_object_media_image(ctx: &mut RtfParseContext) {
    if let Some(obj) = ctx.object_stack.last_mut() {
        obj.media_type = Some(super::MediaType::Image);
    }
}

fn set_last_object_dimension(
    ctx: &mut RtfParseContext,
    param: Option<i32>,
    mut apply: impl FnMut(&mut ObjectContext, u32),
) {
    if let Some(value) = param {
        if let Some(obj) = ctx.object_stack.last_mut() {
            apply(obj, value.max(0) as u32);
        }
    }
}

fn set_pending_cell_vertical_align(ctx: &mut RtfParseContext, value: CellVerticalAlignment) {
    ctx.pending_cell_props
        .get_or_insert_with(TableCellProperties::default)
        .vertical_align = Some(value);
}

fn set_pending_border_style(ctx: &mut RtfParseContext, style: BorderStyle) {
    ctx.pending_border.style = style;
    apply_border(ctx);
}

pub(super) fn handle_group_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
) -> Result<bool, ParseError> {
    match word {
        "fonttbl" => {
            set_last_group_kind(ctx, GroupKind::FontTable);
            ctx.current_text.clear();
        }
        "colortbl" => {
            set_last_group_kind(ctx, GroupKind::ColorTable);
            if ctx.color_table.colors.is_empty() {
                ctx.color_table.colors.push(None);
            }
            ctx.current_text.clear();
        }
        "stylesheet" => {
            set_last_group_kind(ctx, GroupKind::Stylesheet);
            if ctx.style_set.is_none() {
                ctx.style_set = Some(StyleSet::new());
            }
        }
        "s" => {
            if let Some(id) = param {
                let style_id = format!("s{}", id.max(0));
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    set_stylesheet_entry(ctx, style_id, StyleType::Paragraph);
                } else {
                    ctx.pending_para_style = Some(style_id);
                    if let Some(para) = ctx.current_paragraph.as_mut() {
                        para.style_id = ctx.pending_para_style.clone();
                    }
                }
            }
        }
        "cs" => {
            if let Some(id) = param {
                let style_id = format!("cs{}", id.max(0));
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    set_stylesheet_entry(ctx, style_id, StyleType::Character);
                } else {
                    ctx.current_props.style_id = Some(style_id);
                }
            }
        }
        "ds" => {
            if let Some(id) = param {
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    set_stylesheet_entry(ctx, format!("ds{}", id.max(0)), StyleType::Other);
                }
            }
        }
        "ts" => {
            if let Some(id) = param {
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    set_stylesheet_entry(ctx, format!("ts{}", id.max(0)), StyleType::Table);
                }
            }
        }
        "info" => {
            set_last_group_kind(ctx, GroupKind::Info);
        }
        "listtable" => {
            set_last_group_kind(ctx, GroupKind::ListTable);
        }
        "listoverridetable" => {
            set_last_group_kind(ctx, GroupKind::ListOverrideTable);
        }
        "list" => {
            if ctx.current_group_kind() == GroupKind::ListTable {
                set_last_group_kind(ctx, GroupKind::List);
                ctx.current_list_id = None;
                ctx.current_list_level = 0;
            }
        }
        "listlevel" => {
            if ctx.current_group_kind() == GroupKind::List {
                ctx.current_list_level = ctx.current_list_level.saturating_add(1);
                set_last_group_kind(ctx, GroupKind::ListLevel);
            }
        }
        "levelnfc" => {
            if let (Some(list_id), Some(nfc)) = (ctx.current_list_id, param) {
                let level = ctx.current_list_level.saturating_sub(1);
                ctx.list_level_formats
                    .insert((list_id, level), format!("nfc:{}", nfc));
            }
        }
        "listid" => {
            if let Some(id) = param {
                if matches!(
                    ctx.current_group_kind(),
                    GroupKind::List | GroupKind::ListLevel | GroupKind::ListOverride
                ) {
                    ctx.current_list_id = Some(id);
                    if ctx.current_group_kind() == GroupKind::ListOverride {
                        ctx.current_list_override_list_id = Some(id);
                    }
                }
            }
        }
        "listoverride" => {
            if ctx.current_group_kind() == GroupKind::ListOverrideTable {
                set_last_group_kind(ctx, GroupKind::ListOverride);
                ctx.current_list_override = None;
                ctx.current_list_override_list_id = None;
            }
        }
        "ls" => {
            if let Some(id) = param {
                if ctx.current_group_kind() == GroupKind::ListOverride {
                    ctx.current_list_override = Some(id);
                } else {
                    ctx.pending_list_override = Some(id);
                    apply_pending_numbering_to_current_paragraph(ctx);
                }
            }
        }
        "ilvl" => {
            if let Some(level) = param {
                ctx.pending_list_level = Some(level.max(0) as u32);
                apply_pending_numbering_to_current_paragraph(ctx);
            }
        }
        _ => return Ok(false),
    }
    Ok(true)
}

pub(super) fn handle_table_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<bool, ParseError> {
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
        "clvertalt" => {
            set_pending_cell_vertical_align(ctx, CellVerticalAlignment::Top);
        }
        "clvertalc" => {
            set_pending_cell_vertical_align(ctx, CellVerticalAlignment::Center);
        }
        "clvertalb" => {
            set_pending_cell_vertical_align(ctx, CellVerticalAlignment::Bottom);
        }
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

pub(super) fn handle_object_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
) -> bool {
    match word {
        "pict" => {
            set_last_group_kind(ctx, GroupKind::Picture);
            ctx.object_stack.push(ObjectContext::default());
        }
        "pngblip" => {
            set_last_object_media_image(ctx);
        }
        "jpegblip" | "jpgblip" => {
            set_last_object_media_image(ctx);
        }
        "wmetafile" | "emfblip" | "wmetafile8" => {
            set_last_object_media_image(ctx);
        }
        "picw" => {
            set_last_object_dimension(ctx, param, |obj, value| {
                obj.pic_width = Some(value);
            });
        }
        "pich" => {
            set_last_object_dimension(ctx, param, |obj, value| {
                obj.pic_height = Some(value);
            });
        }
        "picwgoal" => {
            set_last_object_dimension(ctx, param, |obj, value| {
                obj.pic_width = Some(value);
            });
        }
        "pichgoal" => {
            set_last_object_dimension(ctx, param, |obj, value| {
                obj.pic_height = Some(value);
            });
        }
        "object" => {
            set_last_group_kind(ctx, GroupKind::Object);
            ctx.object_stack.push(ObjectContext::default());
        }
        "objclass" => {
            ctx.current_text.clear();
            ctx.object_text_target = Some(ObjectTextTarget::Class);
        }
        "objname" => {
            ctx.current_text.clear();
            ctx.object_text_target = Some(ObjectTextTarget::Name);
        }
        "objdata" => {
            set_last_group_kind(ctx, GroupKind::Object);
            ctx.current_text.clear();
            ctx.object_text_target = None;
        }
        _ => return false,
    }
    true
}
