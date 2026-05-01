use super::{
    finalize_object, finalize_picture, GroupKind, IrStore, RtfParseContext, StyleEntryContext,
};
use docir_core::ir::{Style, StyleSet, StyleType};

pub(super) fn handle_group_end(ctx: &mut RtfParseContext, store: &mut IrStore) {
    match ctx.current_group_kind() {
        GroupKind::Field => super::finalize_field(ctx, store),
        GroupKind::Picture => {
            if let Some(obj) = ctx.object_stack.pop() {
                if let Some(asset_id) = finalize_picture(obj, store) {
                    ctx.media_assets.push(asset_id);
                }
            }
        }
        GroupKind::Object => {
            if let Some(obj) = ctx.object_stack.pop() {
                if let Some(ole_id) = finalize_object(obj, store) {
                    ctx.ole_objects.push(ole_id);
                }
            }
        }
        GroupKind::StylesheetEntry => finalize_style_entry(ctx),
        GroupKind::List => finalize_list_entry(ctx),
        GroupKind::ListOverride => finalize_list_override(ctx),
        _ => {}
    }
}

pub(super) fn finalize_style_entry(ctx: &mut RtfParseContext) {
    let Some(style_ctx) = ctx.current_style.take() else {
        return;
    };
    let name = style_ctx.name_buf.trim().to_string();
    if !name.is_empty() {
        push_style_from_ctx(ctx, &style_ctx, name);
    }
}

pub(super) fn push_style_from_ctx(
    ctx: &mut RtfParseContext,
    style_ctx: &StyleEntryContext,
    name: String,
) {
    let style_id = style_ctx
        .style_id
        .clone()
        .unwrap_or_else(|| "style".to_string());
    let style_type = style_ctx.style_type.unwrap_or(StyleType::Other);
    let style = Style {
        style_id,
        name: if name.is_empty() { None } else { Some(name) },
        style_type,
        based_on: None,
        next: None,
        is_default: false,
        run_props: None,
        paragraph_props: None,
        table_props: None,
    };
    if ctx.style_set.is_none() {
        ctx.style_set = Some(StyleSet::new());
    }
    if let Some(set) = ctx.style_set.as_mut() {
        set.styles.push(style);
    }
}

pub(super) fn finalize_list_entry(ctx: &mut RtfParseContext) {
    if let Some(list_id) = ctx.current_list_id {
        let levels = ctx.current_list_level.max(1);
        ctx.list_levels.insert(list_id, levels);
    }
    ctx.current_list_id = None;
    ctx.current_list_level = 0;
}

pub(super) fn finalize_list_override(ctx: &mut RtfParseContext) {
    if let (Some(override_id), Some(list_id)) =
        (ctx.current_list_override, ctx.current_list_override_list_id)
    {
        ctx.list_overrides.insert(override_id, list_id);
    }
    ctx.current_list_override = None;
    ctx.current_list_override_list_id = None;
}
