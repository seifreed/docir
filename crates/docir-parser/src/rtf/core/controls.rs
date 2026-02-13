use super::{
    apply_border, apply_paragraph_border, color_from_index, ensure_section, finalize_cell,
    finalize_row, flush_text, pending_numbering, set_border_target, BorderStyle, BorderTarget,
    CellVerticalAlignment, GroupKind, MergeType, ObjectContext, ObjectTextTarget, RtfParseContext,
    StyleEntryContext, StyleSet, StyleType, Table, TableCell, TableCellProperties, TableRow,
};
use crate::error::ParseError;
use docir_core::visitor::IrStore;

pub(super) fn handle_group_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
) -> Result<bool, ParseError> {
    match word {
        "fonttbl" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::FontTable;
            }
            ctx.current_text.clear();
        }
        "colortbl" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::ColorTable;
            }
            if ctx.color_table.colors.is_empty() {
                ctx.color_table.colors.push(None);
            }
            ctx.current_text.clear();
        }
        "stylesheet" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Stylesheet;
            }
            if ctx.style_set.is_none() {
                ctx.style_set = Some(StyleSet::new());
            }
        }
        "s" => {
            if let Some(id) = param {
                let style_id = format!("s{}", id.max(0));
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    ctx.group_stack
                        .last_mut()
                        .map(|g| g.kind = GroupKind::StylesheetEntry);
                    ctx.current_style = Some(StyleEntryContext {
                        style_id: Some(style_id),
                        style_type: Some(StyleType::Paragraph),
                        name_buf: String::new(),
                    });
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
                    ctx.group_stack
                        .last_mut()
                        .map(|g| g.kind = GroupKind::StylesheetEntry);
                    ctx.current_style = Some(StyleEntryContext {
                        style_id: Some(style_id),
                        style_type: Some(StyleType::Character),
                        name_buf: String::new(),
                    });
                } else {
                    ctx.current_props.style_id = Some(style_id);
                }
            }
        }
        "ds" => {
            if let Some(id) = param {
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    ctx.group_stack
                        .last_mut()
                        .map(|g| g.kind = GroupKind::StylesheetEntry);
                    ctx.current_style = Some(StyleEntryContext {
                        style_id: Some(format!("ds{}", id.max(0))),
                        style_type: Some(StyleType::Other),
                        name_buf: String::new(),
                    });
                }
            }
        }
        "ts" => {
            if let Some(id) = param {
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    ctx.group_stack
                        .last_mut()
                        .map(|g| g.kind = GroupKind::StylesheetEntry);
                    ctx.current_style = Some(StyleEntryContext {
                        style_id: Some(format!("ts{}", id.max(0))),
                        style_type: Some(StyleType::Table),
                        name_buf: String::new(),
                    });
                }
            }
        }
        "info" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Info;
            }
        }
        "listtable" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::ListTable;
            }
        }
        "listoverridetable" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::ListOverrideTable;
            }
        }
        "list" => {
            if ctx.current_group_kind() == GroupKind::ListTable {
                if let Some(group) = ctx.group_stack.last_mut() {
                    group.kind = GroupKind::List;
                }
                ctx.current_list_id = None;
                ctx.current_list_level = 0;
            }
        }
        "listlevel" => {
            if ctx.current_group_kind() == GroupKind::List {
                ctx.current_list_level = ctx.current_list_level.saturating_add(1);
                if let Some(group) = ctx.group_stack.last_mut() {
                    group.kind = GroupKind::ListLevel;
                }
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
                if let Some(group) = ctx.group_stack.last_mut() {
                    group.kind = GroupKind::ListOverride;
                }
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
                    let numbering = pending_numbering(ctx);
                    if let Some(para) = ctx.current_paragraph.as_mut() {
                        if let Some(numbering) = numbering {
                            para.properties.numbering = Some(numbering);
                        }
                    }
                }
            }
        }
        "ilvl" => {
            if let Some(level) = param {
                ctx.pending_list_level = Some(level.max(0) as u32);
                let numbering = pending_numbering(ctx);
                if let Some(para) = ctx.current_paragraph.as_mut() {
                    if let Some(numbering) = numbering {
                        para.properties.numbering = Some(numbering);
                    }
                }
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
            ctx.pending_cell_props
                .get_or_insert_with(TableCellProperties::default)
                .vertical_align = Some(CellVerticalAlignment::Top);
        }
        "clvertalc" => {
            ctx.pending_cell_props
                .get_or_insert_with(TableCellProperties::default)
                .vertical_align = Some(CellVerticalAlignment::Center);
        }
        "clvertalb" => {
            ctx.pending_cell_props
                .get_or_insert_with(TableCellProperties::default)
                .vertical_align = Some(CellVerticalAlignment::Bottom);
        }
        "clbrdrt" => set_border_target(ctx, BorderTarget::Top),
        "clbrdrb" => set_border_target(ctx, BorderTarget::Bottom),
        "clbrdrl" => set_border_target(ctx, BorderTarget::Left),
        "clbrdrr" => set_border_target(ctx, BorderTarget::Right),
        "clbrdrh" => set_border_target(ctx, BorderTarget::InsideH),
        "clbrdrv" => set_border_target(ctx, BorderTarget::InsideV),
        "brdrs" => {
            ctx.pending_border.style = BorderStyle::Single;
            apply_border(ctx);
        }
        "brdrth" => {
            ctx.pending_border.style = BorderStyle::Thick;
            apply_border(ctx);
        }
        "brdrdb" => {
            ctx.pending_border.style = BorderStyle::Double;
            apply_border(ctx);
        }
        "brdrdot" => {
            ctx.pending_border.style = BorderStyle::Dotted;
            apply_border(ctx);
        }
        "brdrdash" => {
            ctx.pending_border.style = BorderStyle::Dashed;
            apply_border(ctx);
        }
        "brdrtriple" => {
            ctx.pending_border.style = BorderStyle::Triple;
            apply_border(ctx);
        }
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
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Picture;
            }
            ctx.object_stack.push(ObjectContext::default());
        }
        "pngblip" => {
            if let Some(obj) = ctx.object_stack.last_mut() {
                obj.media_type = Some(super::MediaType::Image);
            }
        }
        "jpegblip" | "jpgblip" => {
            if let Some(obj) = ctx.object_stack.last_mut() {
                obj.media_type = Some(super::MediaType::Image);
            }
        }
        "wmetafile" | "emfblip" | "wmetafile8" => {
            if let Some(obj) = ctx.object_stack.last_mut() {
                obj.media_type = Some(super::MediaType::Image);
            }
        }
        "picw" => {
            if let Some(value) = param {
                if let Some(obj) = ctx.object_stack.last_mut() {
                    obj.pic_width = Some(value.max(0) as u32);
                }
            }
        }
        "pich" => {
            if let Some(value) = param {
                if let Some(obj) = ctx.object_stack.last_mut() {
                    obj.pic_height = Some(value.max(0) as u32);
                }
            }
        }
        "picwgoal" => {
            if let Some(value) = param {
                if let Some(obj) = ctx.object_stack.last_mut() {
                    obj.pic_width = Some(value.max(0) as u32);
                }
            }
        }
        "pichgoal" => {
            if let Some(value) = param {
                if let Some(obj) = ctx.object_stack.last_mut() {
                    obj.pic_height = Some(value.max(0) as u32);
                }
            }
        }
        "object" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Object;
            }
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
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Object;
            }
            ctx.current_text.clear();
            ctx.object_text_target = None;
        }
        _ => return false,
    }
    true
}
