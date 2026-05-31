use super::super::state::{GroupKind, RtfParseContext};
use super::set_last_group_kind;
use crate::rtf::objects::{ObjectContext, ObjectTextTarget};
use docir_core::ir::MediaType;

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
            flush_pending_object_text(ctx);
            ctx.object_text_target = Some(ObjectTextTarget::Class);
        }
        "objname" => {
            flush_pending_object_text(ctx);
            ctx.object_text_target = Some(ObjectTextTarget::Name);
        }
        "objdata" => {
            flush_pending_object_text(ctx);
            set_last_group_kind(ctx, GroupKind::Object);
            ctx.object_text_target = None;
        }
        _ => return false,
    }
    true
}

fn flush_pending_object_text(ctx: &mut RtfParseContext) {
    let Some(target) = ctx.object_text_target else {
        ctx.current_text.clear();
        return;
    };
    if ctx.current_text.is_empty() {
        return;
    }
    let text = std::mem::take(&mut ctx.current_text);
    let Some(obj) = ctx.object_stack.last_mut() else {
        return;
    };
    match target {
        ObjectTextTarget::Class => obj.class_name = Some(text.trim().to_string()),
        ObjectTextTarget::Name => obj.object_name = Some(text.trim().to_string()),
    }
}

fn set_last_object_media_image(ctx: &mut RtfParseContext) {
    if let Some(obj) = ctx.object_stack.last_mut() {
        obj.media_type = Some(MediaType::Image);
    }
}

fn set_last_object_dimension(
    ctx: &mut RtfParseContext,
    param: Option<i32>,
    mut apply: impl FnMut(&mut ObjectContext, u32),
) {
    if let Some(value) = param
        && let Some(obj) = ctx.object_stack.last_mut()
    {
        apply(obj, value.max(0) as u32);
    }
}
