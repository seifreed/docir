//! Event/state helpers for RTF parse orchestration.

use crate::error::ParseError;
use docir_core::ir::{
    Paragraph, ParagraphBorders, Run, RunProperties, TableCellProperties, TableWidth,
    TableWidthType,
};
use docir_core::types::SourceSpan;
use docir_core::visitor::IrStore;

use super::super::controls::{
    handle_group_controls, handle_object_controls, handle_table_controls,
};
use super::super::cursor::RtfCursor;
use super::super::helpers::{
    attach_flushed_run, flush_object_text, flush_stylesheet_text, handle_encoding_controls,
    handle_paragraph_controls, hex_val, parse_color_entries, parse_font_entry,
};
use super::super::state::{BorderTarget, FieldContext, GroupKind, RtfParseContext};

pub(crate) fn parse_control(
    cursor: &mut RtfCursor<'_>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<(), ParseError> {
    let Some(next) = cursor.next() else {
        return Ok(());
    };

    if try_handle_simple_control(cursor, ctx, next)? {
        return Ok(());
    }
    if !next.is_ascii_alphabetic() {
        return Ok(());
    }

    let (word, param) = parse_control_word_and_param(cursor, next);
    if handle_control_word_with_guard(&word, param, ctx, store)? {
        return Ok(());
    }
    Ok(())
}

fn try_handle_simple_control(
    cursor: &mut RtfCursor<'_>,
    ctx: &mut RtfParseContext,
    next: u8,
) -> Result<bool, ParseError> {
    match next {
        b'\\' | b'{' | b'}' => {
            append_text_byte(ctx, next);
            return Ok(true);
        }
        b'\'' => {
            let Some(hi) = cursor.next() else {
                return Ok(true);
            };
            let Some(lo) = cursor.next() else {
                return Ok(true);
            };
            if let (Some(h), Some(l)) = (hex_val(hi), hex_val(lo)) {
                append_text_byte(ctx, (h << 4) | l);
            }
            return Ok(true);
        }
        b'*' => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Skip;
            }
            return Ok(true);
        }
        b'~' => {
            append_text(ctx, " ");
            return Ok(true);
        }
        b'-' => {
            return Ok(true); // optional hyphen
        }
        b'_' => {
            append_text(ctx, "-");
            return Ok(true);
        }
        b'\n' | b'\r' => return Ok(true),
        _ => {}
    }
    Ok(false)
}

pub(crate) fn parse_control_word_and_param(
    cursor: &mut RtfCursor<'_>,
    first: u8,
) -> (String, Option<i32>) {
    let mut word = vec![first];
    while let Some(b) = cursor.peek() {
        if b.is_ascii_alphabetic() {
            word.push(b);
            cursor.next();
        } else {
            break;
        }
    }

    let mut sign = 1i32;
    if cursor.peek() == Some(b'-') {
        sign = -1;
        cursor.next();
    }

    let mut digits = Vec::new();
    while let Some(b) = cursor.peek() {
        if b.is_ascii_digit() && digits.len() < 10 {
            digits.push(b);
            cursor.next();
        } else {
            break;
        }
    }

    let param = if digits.is_empty() {
        None
    } else {
        std::str::from_utf8(&digits)
            .ok()
            .and_then(|raw| raw.parse::<i32>().ok())
            .map(|num| num * sign)
    };

    if cursor.peek() == Some(b' ') {
        cursor.next();
    }

    (String::from_utf8_lossy(&word).to_ascii_lowercase(), param)
}

fn handle_control_word_with_guard(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<bool, ParseError> {
    if handle_paragraph_controls(word, param, ctx, store)? {
        return Ok(true);
    }
    if handle_run_style_controls(word, param, ctx) {
        return Ok(true);
    }
    if handle_group_controls(word, param, ctx)? {
        return Ok(true);
    }
    if handle_field_controls(word, ctx) {
        return Ok(true);
    }
    if handle_object_controls(word, param, ctx) {
        return Ok(true);
    }
    if handle_table_controls(word, param, ctx, store)? {
        return Ok(true);
    }
    if handle_encoding_controls(word, param, ctx) {
        return Ok(true);
    }
    Ok(false)
}

fn handle_field_controls(word: &str, ctx: &mut RtfParseContext) -> bool {
    match word {
        "field" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Field;
            }
            ctx.field_stack.push(FieldContext::default());
        }
        "fldinst" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::FieldInst;
            }
            ctx.current_text.clear();
        }
        "fldrslt" => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::FieldResult;
            }
            ctx.current_text.clear();
        }
        _ => return false,
    }
    true
}

fn handle_run_style_controls(word: &str, param: Option<i32>, ctx: &mut RtfParseContext) -> bool {
    match word {
        "b" => {
            let param = param.unwrap_or(1);
            ctx.current_props.bold = Some(param != 0);
        }
        "i" => {
            let param = param.unwrap_or(1);
            ctx.current_props.italic = Some(param != 0);
        }
        "ul" => {
            let param = param.unwrap_or(1);
            ctx.current_props.underline = Some(param != 0);
        }
        "ulnone" => {
            ctx.current_props.underline = Some(false);
        }
        "strike" => {
            let param = param.unwrap_or(1);
            ctx.current_props.strike = Some(param != 0);
        }
        "fs" => {
            if let Some(sz) = param {
                ctx.current_props.font_size = Some(sz.max(0) as u32);
            }
        }
        "f" => {
            if let Some(idx) = param {
                ctx.current_props.font_index = Some(idx.max(0) as u32);
            }
        }
        "cf" => {
            if let Some(idx) = param {
                ctx.current_props.color_index = Some(idx.max(0) as usize);
            }
        }
        "highlight" => {
            if let Some(idx) = param {
                ctx.current_props.highlight_index = Some(idx.max(0) as usize);
            }
        }
        "super" => {
            ctx.current_props.vertical = Some(docir_core::ir::VerticalTextAlignment::Superscript);
        }
        "sub" => {
            ctx.current_props.vertical = Some(docir_core::ir::VerticalTextAlignment::Subscript);
        }
        "nosupersub" => {
            ctx.current_props.vertical = Some(docir_core::ir::VerticalTextAlignment::Baseline);
        }
        _ => return false,
    }
    true
}

pub(crate) fn flush_text(
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
    span: Option<SourceSpan>,
) -> Result<(), ParseError> {
    if ctx.current_text.is_empty() {
        return Ok(());
    }

    let text = ctx.current_text.clone();
    ctx.current_text.clear();

    if ctx.current_group_kind() == GroupKind::FontTable {
        parse_font_entry(&text, ctx);
        return Ok(());
    }
    if ctx.current_group_kind() == GroupKind::ColorTable {
        parse_color_entries(&text, ctx);
        return Ok(());
    }
    if matches!(
        ctx.current_group_kind(),
        GroupKind::Stylesheet | GroupKind::StylesheetEntry
    ) {
        flush_stylesheet_text(ctx, &text);
        return Ok(());
    }
    if ctx.current_group_kind() == GroupKind::Object {
        flush_object_text(ctx, &text);
        return Ok(());
    }

    let props = run_properties_from_state(ctx);
    let mut run = Run::with_properties(text.clone(), props);
    run.span = span.clone();
    let run_id = run.id;
    store.insert(docir_core::ir::IRNode::Run(run));

    attach_flushed_run(ctx, store, &text, run_id);

    Ok(())
}

pub(crate) fn append_text(ctx: &mut RtfParseContext, text: &str) {
    ctx.current_text.push_str(text);
}

pub(crate) fn append_text_byte(ctx: &mut RtfParseContext, byte: u8) {
    let binding = [byte];
    let (text, _, _) = ctx.encoding.decode(&binding);
    ctx.current_text.push_str(&text);
}

pub(crate) fn ensure_paragraph(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if ctx.current_paragraph.is_none() {
        let mut para = Paragraph::new();
        para.span = Some(SourceSpan::new("rtf"));
        apply_pending_paragraph(&mut para, ctx);
        ctx.current_paragraph = Some(para);
    }
    ensure_section(ctx, store);
}

pub(crate) fn ensure_section(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if ctx.current_section.is_none() {
        let mut section = docir_core::ir::Section::new();
        section.span = Some(SourceSpan::new("rtf"));
        let section_id = section.id;
        store.insert(docir_core::ir::IRNode::Section(section));
        ctx.current_section = Some(section_id);
        ctx.sections.push(section_id);
    }
}

fn apply_pending_paragraph(para: &mut Paragraph, ctx: &mut RtfParseContext) {
    if let Some(style_id) = ctx.pending_para_style.clone() {
        para.style_id = Some(style_id);
    }
    if let Some(numbering) = pending_numbering(ctx) {
        para.properties.numbering = Some(numbering);
    }
    if let Some(align) = ctx.pending_alignment {
        para.properties.alignment = Some(align);
    }
    if ctx.pending_indent.left.is_some()
        || ctx.pending_indent.right.is_some()
        || ctx.pending_indent.first_line.is_some()
        || ctx.pending_indent.hanging.is_some()
    {
        para.properties.indentation = Some(ctx.pending_indent.clone());
    }
    if ctx.pending_spacing.before.is_some()
        || ctx.pending_spacing.after.is_some()
        || ctx.pending_spacing.line.is_some()
        || ctx.pending_spacing.line_rule.is_some()
    {
        para.properties.spacing = Some(ctx.pending_spacing.clone());
    }
    if let Some(borders) = pending_paragraph_borders(ctx) {
        para.properties.borders = Some(borders);
    }
}

pub(crate) fn pending_numbering(ctx: &RtfParseContext) -> Option<docir_core::ir::NumberingInfo> {
    let list_override = ctx.pending_list_override?;
    let level = ctx.pending_list_level.unwrap_or(0);
    let list_id = ctx
        .list_overrides
        .get(&list_override)
        .copied()
        .unwrap_or(list_override);
    let format = ctx.list_level_formats.get(&(list_id, level)).cloned();
    Some(docir_core::ir::NumberingInfo {
        num_id: list_id.max(0) as u32,
        level,
        format,
    })
}

pub(crate) fn color_from_index(ctx: &RtfParseContext, index: usize) -> Option<String> {
    ctx.color_table.colors.get(index).and_then(|c| c.clone())
}

pub(crate) fn set_border_target(ctx: &mut RtfParseContext, target: BorderTarget) {
    ctx.pending_border_target = Some(target);
}

pub(crate) fn apply_border(ctx: &mut RtfParseContext) {
    if ctx.pending_para_border_target.is_some() {
        ctx.pending_para_border = ctx.pending_border.clone();
        apply_paragraph_border(ctx);
        return;
    }
    let Some(target) = ctx.pending_border_target else {
        return;
    };
    let props = ctx
        .pending_cell_props
        .get_or_insert_with(TableCellProperties::default);
    let mut borders = props.borders.take().unwrap_or_default();
    let border = ctx.pending_border.clone();
    match target {
        BorderTarget::Top => borders.top = Some(border),
        BorderTarget::Bottom => borders.bottom = Some(border),
        BorderTarget::Left => borders.left = Some(border),
        BorderTarget::Right => borders.right = Some(border),
        BorderTarget::InsideH => borders.inside_h = Some(border),
        BorderTarget::InsideV => borders.inside_v = Some(border),
    }
    props.borders = Some(borders);
}

pub(crate) fn apply_paragraph_border(ctx: &mut RtfParseContext) {
    let Some(target) = ctx.pending_para_border_target else {
        return;
    };
    let border = ctx.pending_para_border.clone();
    if let Some(para) = ctx.current_paragraph.as_mut() {
        let mut borders = para.properties.borders.take().unwrap_or_default();
        match target {
            BorderTarget::Top => borders.top = Some(border),
            BorderTarget::Bottom => borders.bottom = Some(border),
            BorderTarget::Left => borders.left = Some(border),
            BorderTarget::Right => borders.right = Some(border),
            BorderTarget::InsideH | BorderTarget::InsideV => {}
        }
        para.properties.borders = Some(borders);
    } else {
        match target {
            BorderTarget::Top => ctx.pending_para_borders.top = Some(border),
            BorderTarget::Bottom => ctx.pending_para_borders.bottom = Some(border),
            BorderTarget::Left => ctx.pending_para_borders.left = Some(border),
            BorderTarget::Right => ctx.pending_para_borders.right = Some(border),
            BorderTarget::InsideH | BorderTarget::InsideV => {}
        }
    }
}

fn pending_paragraph_borders(ctx: &RtfParseContext) -> Option<ParagraphBorders> {
    let borders = &ctx.pending_para_borders;
    if borders.top.is_none()
        && borders.bottom.is_none()
        && borders.left.is_none()
        && borders.right.is_none()
    {
        None
    } else {
        Some(borders.clone())
    }
}

pub(crate) fn cell_width_from_row(ctx: &RtfParseContext) -> Option<TableWidth> {
    if ctx.row_cellx.is_empty() {
        return None;
    }
    let idx = ctx.current_cell_index;
    let end = *ctx.row_cellx.get(idx)?;
    let start = if idx == 0 {
        0
    } else {
        *ctx.row_cellx.get(idx - 1)?
    };
    let width = (end - start).max(0) as u32;
    Some(TableWidth {
        value: width,
        width_type: TableWidthType::Dxa,
    })
}

pub(crate) fn run_properties_from_state(ctx: &RtfParseContext) -> RunProperties {
    let mut props = RunProperties {
        style_id: ctx.current_props.style_id.clone(),
        bold: ctx.current_props.bold,
        italic: ctx.current_props.italic,
        underline: ctx.current_props.underline.map(|u| {
            if u {
                docir_core::ir::UnderlineStyle::Single
            } else {
                docir_core::ir::UnderlineStyle::None
            }
        }),
        strike: ctx.current_props.strike,
        font_size: ctx.current_props.font_size,
        vertical_align: ctx.current_props.vertical,
        ..RunProperties::default()
    };
    if let Some(idx) = ctx.current_props.font_index {
        if let Some(name) = ctx.font_table.fonts.get(&idx) {
            props.font_family = Some(name.clone());
        }
    }
    if let Some(idx) = ctx.current_props.color_index {
        if let Some(color) = ctx.color_table.colors.get(idx).and_then(|c| c.clone()) {
            props.color = Some(color);
        }
    }
    if let Some(idx) = ctx.current_props.highlight_index {
        if let Some(color) = ctx.color_table.colors.get(idx).and_then(|c| c.clone()) {
            props.highlight = Some(color);
        }
    }
    props
}
