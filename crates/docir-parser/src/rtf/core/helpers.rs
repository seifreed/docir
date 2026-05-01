use crate::error::ParseError;
use docir_core::ir::{Indentation, LineSpacingRule, ParagraphBorders, Spacing, TextAlignment};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;
use encoding_rs::Encoding;

use super::super::objects::ObjectTextTarget;
use super::core_parse::push_style_from_ctx;
use super::state::RtfStyleState;
use super::state::{GroupKind, RtfParseContext};
use super::{append_text, ensure_paragraph, finalize_paragraph, flush_text};

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
                ctx.pending_spacing.line = Some(value.unsigned_abs());
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
    ctx.pending_alignment = Some(alignment);
    if let Some(para) = ctx.current_paragraph.as_mut() {
        para.properties.alignment = Some(alignment);
    }
}

pub(super) fn parse_font_entry(text: &str, ctx: &mut RtfParseContext) {
    let header = text.split(';').next().unwrap_or(text);

    let mut font_index = None;
    for_each_control_token(header, |token| {
        if let Some(num) = token.strip_prefix('f') {
            if let Ok(idx) = num.parse::<u32>() {
                font_index = Some(idx);
            }
        }
    });

    let name = parse_control_free_text(header).trim().to_string();
    if let Some(idx) = font_index {
        if !name.is_empty() {
            ctx.font_table.fonts.insert(idx, name);
        }
    }
}

pub(super) fn parse_color_entries(text: &str, ctx: &mut RtfParseContext) {
    let mut rest = text;
    while let Some((chunk, tail)) = rest.split_once(';') {
        if let Some(color) = parse_color_chunk(chunk) {
            ctx.color_table.colors.push(Some(color));
        } else {
            ctx.color_table.colors.push(None);
        }
        rest = tail;
    }
}

fn parse_color_chunk(chunk: &str) -> Option<String> {
    let mut red = None;
    let mut green = None;
    let mut blue = None;

    for_each_control_token(chunk, |token| {
        if token.starts_with('r') {
            red = token.strip_prefix('r').and_then(|v| v.parse::<u8>().ok());
        } else if token.starts_with('g') {
            green = token.strip_prefix('g').and_then(|v| v.parse::<u8>().ok());
        } else if token.starts_with('b') {
            blue = token.strip_prefix('b').and_then(|v| v.parse::<u8>().ok());
        }
    });

    match (red, green, blue) {
        (Some(r), Some(g), Some(b)) => Some(format!("{:02X}{:02X}{:02X}", r, g, b)),
        _ => None,
    }
}

fn parse_control_free_text(text: &str) -> String {
    let bytes = text.as_bytes();
    let mut out = String::new();
    let mut idx = 0usize;
    while idx < bytes.len() {
        if bytes[idx] == b'\\' {
            idx += 1;
            while idx < bytes.len() && is_alpha_num_ascii(bytes[idx]) {
                idx += 1;
            }
            continue;
        }
        out.push(bytes[idx] as char);
        idx += 1;
    }
    out
}

fn for_each_control_token<F>(text: &str, mut on_token: F)
where
    F: FnMut(&str),
{
    let bytes = text.as_bytes();
    let mut idx = 0usize;
    while idx < bytes.len() {
        if bytes[idx] != b'\\' {
            idx += 1;
            continue;
        }
        idx += 1;
        let start = idx;
        while idx < bytes.len() && is_alpha_num_ascii(bytes[idx]) {
            idx += 1;
        }
        if start < idx {
            if let Ok(token) = std::str::from_utf8(&bytes[start..idx]) {
                on_token(token);
            }
        }
    }
}

fn is_alpha_num_ascii(byte: u8) -> bool {
    byte.is_ascii_alphanumeric()
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::rtf::core::state::GroupState;
    use docir_core::ir::Paragraph;
    use docir_core::visitor::IrStore;

    fn ctx() -> RtfParseContext {
        RtfParseContext::new(128, 0)
    }

    #[test]
    fn handle_encoding_controls_sets_known_codepage() {
        let mut ctx = ctx();
        let handled = handle_encoding_controls("ansicpg", Some(1251), &mut ctx);
        assert!(handled);
        assert_eq!(ctx.encoding.name(), "windows-1251");
        assert!(!handle_encoding_controls("unknown", None, &mut ctx));
    }

    #[test]
    fn parse_font_entry_records_font_name_and_index() {
        let mut ctx = ctx();
        parse_font_entry(r"\f0\fnil Helvetica;", &mut ctx);
        parse_font_entry(r"\f12 Calibri;", &mut ctx);

        assert_eq!(
            ctx.font_table.fonts.get(&0).map(String::as_str),
            Some("Helvetica")
        );
        assert_eq!(
            ctx.font_table.fonts.get(&12).map(String::as_str),
            Some("Calibri")
        );
    }

    #[test]
    fn parse_color_entries_records_rgb_triplets() {
        let mut ctx = ctx();
        parse_color_entries(r"\r255\g0\b16;\r1\g2\b3;", &mut ctx);

        assert_eq!(ctx.color_table.colors.len(), 2);
        assert_eq!(ctx.color_table.colors[0].as_deref(), Some("FF0010"));
        assert_eq!(ctx.color_table.colors[1].as_deref(), Some("010203"));
    }

    #[test]
    fn flush_object_text_updates_object_fields() {
        let mut ctx = ctx();
        ctx.object_stack
            .push(crate::rtf::objects::ObjectContext::default());

        ctx.object_text_target = Some(ObjectTextTarget::Class);
        flush_object_text(&mut ctx, "  Excel.Sheet.12 ");
        assert_eq!(
            ctx.object_stack[0].class_name.as_deref(),
            Some("Excel.Sheet.12")
        );

        ctx.object_text_target = Some(ObjectTextTarget::Name);
        flush_object_text(&mut ctx, "  Embedded chart ");
        assert_eq!(
            ctx.object_stack[0].object_name.as_deref(),
            Some("Embedded chart")
        );
    }

    #[test]
    fn hex_val_parses_numeric_and_hex_digits() {
        assert_eq!(hex_val(b'0'), Some(0));
        assert_eq!(hex_val(b'9'), Some(9));
        assert_eq!(hex_val(b'a'), Some(10));
        assert_eq!(hex_val(b'f'), Some(15));
        assert_eq!(hex_val(b'A'), Some(10));
        assert_eq!(hex_val(b'F'), Some(15));
        assert_eq!(hex_val(b'z'), None);
    }

    #[test]
    fn handle_paragraph_controls_updates_pending_properties() {
        let mut ctx = ctx();
        let mut store = IrStore::new();
        ctx.current_paragraph = Some(Paragraph::new());

        assert!(handle_paragraph_controls("qc", None, &mut ctx, &mut store).expect("qc"));
        assert!(handle_paragraph_controls("li", Some(720), &mut ctx, &mut store).expect("li"));
        assert!(handle_paragraph_controls("ri", Some(360), &mut ctx, &mut store).expect("ri"));
        assert!(handle_paragraph_controls("fi", Some(-180), &mut ctx, &mut store).expect("fi"));
        assert!(handle_paragraph_controls("sb", Some(120), &mut ctx, &mut store).expect("sb"));
        assert!(handle_paragraph_controls("sa", Some(80), &mut ctx, &mut store).expect("sa"));
        assert!(
            handle_paragraph_controls("slmult", Some(1), &mut ctx, &mut store).expect("slmult")
        );
        assert!(handle_paragraph_controls("sl", Some(240), &mut ctx, &mut store).expect("sl"));
        assert!(
            !handle_paragraph_controls("not_a_control", None, &mut ctx, &mut store)
                .expect("unknown")
        );

        let para = ctx.current_paragraph.as_ref().expect("paragraph");
        assert_eq!(ctx.pending_alignment, Some(TextAlignment::Center));
        assert_eq!(ctx.pending_indent.left, Some(720));
        assert_eq!(ctx.pending_indent.right, Some(360));
        assert_eq!(ctx.pending_indent.first_line, Some(-180));
        assert_eq!(ctx.pending_spacing.before, Some(120));
        assert_eq!(ctx.pending_spacing.after, Some(80));
        assert_eq!(ctx.pending_spacing.line, Some(240));
        assert_eq!(
            ctx.pending_spacing.line_rule,
            Some(LineSpacingRule::AtLeast)
        );
        assert_eq!(para.properties.alignment, Some(TextAlignment::Center));
        let para_indent = para
            .properties
            .indentation
            .as_ref()
            .expect("paragraph indentation");
        assert_eq!(para_indent.left, ctx.pending_indent.left);
        assert_eq!(para_indent.right, ctx.pending_indent.right);
        assert_eq!(para_indent.first_line, ctx.pending_indent.first_line);

        let para_spacing = para.properties.spacing.as_ref().expect("paragraph spacing");
        assert_eq!(para_spacing.before, ctx.pending_spacing.before);
        assert_eq!(para_spacing.after, ctx.pending_spacing.after);
        assert_eq!(para_spacing.line, ctx.pending_spacing.line);
        assert_eq!(para_spacing.line_rule, ctx.pending_spacing.line_rule);
    }

    #[test]
    fn flush_stylesheet_text_accumulates_until_semicolon() {
        let mut ctx = ctx();
        ctx.current_style = Some(crate::rtf::core::state::StyleEntryContext {
            style_id: Some("s1".to_string()),
            style_type: Some(docir_core::ir::StyleType::Paragraph),
            name_buf: "Heading ".to_string(),
        });

        flush_stylesheet_text(&mut ctx, "One");
        assert!(ctx.current_style.is_some());

        flush_stylesheet_text(&mut ctx, ";Trailing");
        assert!(ctx.current_style.is_some());
        assert_eq!(
            ctx.current_style
                .as_ref()
                .map(|style| style.name_buf.as_str()),
            Some("Trailing")
        );

        // Keep stack consistent for context operations used by helper functions.
        assert_eq!(ctx.current_group_kind(), GroupKind::Normal);
        assert_eq!(
            ctx.group_stack.first().map(|g| g.kind),
            Some(GroupKind::Normal)
        );
        let _ = GroupState {
            kind: GroupKind::Normal,
            style: RtfStyleState::default(),
        };
    }
}
