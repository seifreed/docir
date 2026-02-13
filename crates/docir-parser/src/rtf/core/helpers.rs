use crate::error::ParseError;
use docir_core::ir::{Indentation, LineSpacingRule, ParagraphBorders, Spacing, TextAlignment};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;
use encoding_rs::Encoding;

use super::super::objects::ObjectTextTarget;
use super::state::RtfStyleState;
use super::{append_text, ensure_paragraph, finalize_paragraph, flush_text, push_style_from_ctx};
use super::{GroupKind, RtfParseContext};

pub(super) fn handle_encoding_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
) -> bool {
    match word {
        "ansicpg" => {
            if let Some(cp) = param {
                if let Some(enc) = encoding_for_codepage(cp as u32) {
                    ctx.encoding = enc;
                }
            }
        }
        _ => return false,
    }
    true
}

pub(super) fn handle_paragraph_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<bool, ParseError> {
    match word {
        "par" => {
            flush_text(ctx, store, None)?;
            if ctx.current_group_kind() == GroupKind::Normal {
                finalize_paragraph(ctx, store);
            }
        }
        "pard" => {
            flush_text(ctx, store, None)?;
            finalize_paragraph(ctx, store);
            ctx.pending_para_style = None;
            ctx.pending_alignment = None;
            ctx.pending_indent = Indentation::default();
            ctx.pending_spacing = Spacing::default();
            ctx.pending_line_rule = None;
            ctx.pending_para_border_target = None;
            ctx.pending_para_borders = ParagraphBorders::default();
        }
        "plain" => {
            ctx.current_props = RtfStyleState::default();
        }
        "line" => {
            append_text(ctx, "\n");
        }
        "tab" => {
            append_text(ctx, "\t");
        }
        "ql" => {
            apply_paragraph_alignment(ctx, TextAlignment::Left);
        }
        "qr" => {
            apply_paragraph_alignment(ctx, TextAlignment::Right);
        }
        "qc" => {
            apply_paragraph_alignment(ctx, TextAlignment::Center);
        }
        "qj" => {
            apply_paragraph_alignment(ctx, TextAlignment::Justify);
        }
        "li" => {
            if let Some(value) = param {
                apply_pending_indent(ctx, value, |indent, v| indent.left = Some(v));
            }
        }
        "ri" => {
            if let Some(value) = param {
                apply_pending_indent(ctx, value, |indent, v| indent.right = Some(v));
            }
        }
        "fi" => {
            if let Some(value) = param {
                apply_pending_indent(ctx, value, |indent, v| indent.first_line = Some(v));
            }
        }
        "sb" => {
            if let Some(value) = param {
                apply_pending_spacing(ctx, |spacing| spacing.before = Some(value.max(0) as u32));
            }
        }
        "sa" => {
            if let Some(value) = param {
                apply_pending_spacing(ctx, |spacing| spacing.after = Some(value.max(0) as u32));
            }
        }
        "sl" => {
            if let Some(value) = param {
                ctx.pending_spacing.line = Some(value.abs() as u32);
                ctx.pending_spacing.line_rule = ctx.pending_line_rule;
                sync_current_paragraph_spacing(ctx);
            }
        }
        "slmult" => {
            if let Some(value) = param {
                ctx.pending_line_rule = if value == 0 {
                    Some(LineSpacingRule::Exact)
                } else if value == 1 {
                    Some(LineSpacingRule::AtLeast)
                } else {
                    Some(LineSpacingRule::Auto)
                };
            }
        }
        _ => return Ok(false),
    }
    Ok(true)
}

fn apply_pending_indent(
    ctx: &mut RtfParseContext,
    value: i32,
    update: impl FnOnce(&mut Indentation, i32),
) {
    update(&mut ctx.pending_indent, value);
    if let Some(para) = ctx.current_paragraph.as_mut() {
        para.properties.indentation = Some(ctx.pending_indent.clone());
    }
}

fn apply_pending_spacing(ctx: &mut RtfParseContext, update: impl FnOnce(&mut Spacing)) {
    update(&mut ctx.pending_spacing);
    sync_current_paragraph_spacing(ctx);
}

fn sync_current_paragraph_spacing(ctx: &mut RtfParseContext) {
    if let Some(para) = ctx.current_paragraph.as_mut() {
        para.properties.spacing = Some(ctx.pending_spacing.clone());
    }
}

fn apply_paragraph_alignment(ctx: &mut RtfParseContext, alignment: TextAlignment) {
    ctx.pending_alignment = Some(alignment.clone());
    if let Some(para) = ctx.current_paragraph.as_mut() {
        para.properties.alignment = Some(alignment);
    }
}

pub(super) fn parse_font_entry(text: &str, ctx: &mut RtfParseContext) {
    // Parse entries like "\\f0\\fnil Helvetica;"
    let mut font_index: Option<u32> = None;
    let mut name = String::new();
    let mut chars = text.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '\\' {
            let mut control = String::new();
            while let Some(c) = chars.peek() {
                if c.is_alphanumeric() {
                    control.push(*c);
                    chars.next();
                } else {
                    break;
                }
            }
            if control.starts_with('f') {
                let num = control.trim_start_matches('f');
                if let Ok(idx) = num.parse::<u32>() {
                    font_index = Some(idx);
                }
            }
        } else if ch == ';' {
            break;
        } else if !ch.is_control() {
            name.push(ch);
        }
    }
    let name = name.trim().to_string();
    if let Some(idx) = font_index {
        if !name.is_empty() {
            ctx.font_table.fonts.insert(idx, name);
        }
    }
}

pub(super) fn parse_color_entries(text: &str, ctx: &mut RtfParseContext) {
    let mut red: Option<u8> = None;
    let mut green: Option<u8> = None;
    let mut blue: Option<u8> = None;
    let mut chars = text.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '\\' {
            let mut control = String::new();
            while let Some(c) = chars.peek() {
                if c.is_alphanumeric() {
                    control.push(*c);
                    chars.next();
                } else {
                    break;
                }
            }
            if control.starts_with('r') {
                red = control[1..].parse::<u8>().ok();
            } else if control.starts_with('g') {
                green = control[1..].parse::<u8>().ok();
            } else if control.starts_with('b') {
                blue = control[1..].parse::<u8>().ok();
            }
        } else if ch == ';' {
            let color = match (red, green, blue) {
                (Some(r), Some(g), Some(b)) => Some(format!("{:02X}{:02X}{:02X}", r, g, b)),
                _ => None,
            };
            ctx.color_table.colors.push(color);
            red = None;
            green = None;
            blue = None;
        }
    }
}

pub(super) fn flush_stylesheet_text(ctx: &mut RtfParseContext, text: &str) {
    if let Some(mut style_ctx) = ctx.current_style.take() {
        let mut pending = format!("{}{}", style_ctx.name_buf, text);
        loop {
            if let Some(pos) = pending.find(';') {
                let (head, rest) = pending.split_at(pos);
                let name = head.trim().to_string();
                push_style_from_ctx(ctx, &style_ctx, name);
                pending = rest.trim_start_matches(';').to_string();
                style_ctx.name_buf.clear();
            } else {
                style_ctx.name_buf.push_str(pending.trim());
                ctx.current_style = Some(style_ctx);
                break;
            }
        }
    }
}

pub(super) fn flush_object_text(ctx: &mut RtfParseContext, text: &str) {
    let Some(target) = ctx.object_text_target else {
        return;
    };
    let Some(obj) = ctx.object_stack.last_mut() else {
        return;
    };
    match target {
        ObjectTextTarget::Class => obj.class_name = Some(text.trim().to_string()),
        ObjectTextTarget::Name => obj.object_name = Some(text.trim().to_string()),
    }
}

pub(super) fn attach_flushed_run(
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
    text: &str,
    run_id: NodeId,
) {
    match ctx.current_group_kind() {
        GroupKind::FieldInst => {
            if let Some(field) = ctx.field_stack.last_mut() {
                field.instruction.push_str(text);
            }
        }
        GroupKind::FieldResult => {
            if let Some(field) = ctx.field_stack.last_mut() {
                field.runs.push(run_id);
            }
        }
        _ => {
            ensure_paragraph(ctx, store);
            if let Some(para) = ctx.current_paragraph.as_mut() {
                para.runs.push(run_id);
            }
        }
    }
}

pub(super) fn hex_val(b: u8) -> Option<u8> {
    match b {
        b'0'..=b'9' => Some(b - b'0'),
        b'a'..=b'f' => Some(10 + (b - b'a')),
        b'A'..=b'F' => Some(10 + (b - b'A')),
        _ => None,
    }
}

fn encoding_for_codepage(cp: u32) -> Option<&'static Encoding> {
    match cp {
        65001 => Some(encoding_rs::UTF_8),
        1250 => Some(encoding_rs::WINDOWS_1250),
        1251 => Some(encoding_rs::WINDOWS_1251),
        1252 => Some(encoding_rs::WINDOWS_1252),
        1253 => Some(encoding_rs::WINDOWS_1253),
        1254 => Some(encoding_rs::WINDOWS_1254),
        1255 => Some(encoding_rs::WINDOWS_1255),
        1256 => Some(encoding_rs::WINDOWS_1256),
        1257 => Some(encoding_rs::WINDOWS_1257),
        1258 => Some(encoding_rs::WINDOWS_1258),
        _ => None,
    }
}
