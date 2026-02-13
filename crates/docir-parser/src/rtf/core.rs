//! RTF parsing support.

use crate::error::ParseError;
use docir_core::ir::ParagraphBorders;
use docir_core::ir::{
    Border, BorderStyle, CellVerticalAlignment, Indentation, LineSpacingRule, MergeType, Spacing,
    TextAlignment,
};
use docir_core::ir::{
    Field, FieldInstruction, FieldKind, Hyperlink, IRNode, MediaType, NumberingInfo, Paragraph,
    Run, RunProperties, Style, StyleSet, StyleType, Table, TableCell, TableCellProperties,
    TableRow, TableWidth, TableWidthType,
};
use docir_core::security::{ExternalRefType, ExternalReference};
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use encoding_rs::Encoding;

mod controls;
mod cursor;
mod field_utils;
mod helpers;
mod state;

pub(crate) use cursor::{is_rtf_bytes, RtfCursor};
pub(crate) use state::RtfParseContext;
use state::{BorderTarget, FieldContext, GroupKind, StyleEntryContext};

use super::objects::{finalize_object, finalize_picture, ObjectContext, ObjectTextTarget};
use controls::{handle_group_controls, handle_object_controls, handle_table_controls};
use field_utils::{
    parse_field_instruction as parse_field_instruction_impl,
    parse_hyperlink_instruction as parse_hyperlink_instruction_impl,
    tokenize_field_instruction as tokenize_field_instruction_impl,
};
use helpers::{
    handle_encoding_controls, handle_paragraph_controls, hex_val, parse_color_entries,
    parse_font_entry,
};

pub(crate) fn parse_rtf(
    cursor: &mut RtfCursor<'_>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<(), ParseError> {
    while !cursor.is_eof() {
        match cursor.next() {
            Some(b'{') => {
                ctx.push_group(GroupKind::Normal)?;
            }
            Some(b'}') => {
                flush_text(ctx, store, None)?;
                match ctx.current_group_kind() {
                    GroupKind::Field => finalize_field(ctx, store),
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
                    GroupKind::StylesheetEntry => {
                        finalize_style_entry(ctx);
                    }
                    GroupKind::List => {
                        finalize_list_entry(ctx);
                    }
                    GroupKind::ListOverride => {
                        finalize_list_override(ctx);
                    }
                    _ => {}
                }
                ctx.pop_group();
            }
            Some(b'\\') => {
                parse_control(cursor, ctx, store)?;
            }
            Some(b'\r') | Some(b'\n') => {
                // ignore raw newlines
            }
            Some(byte) => match ctx.current_group_kind() {
                GroupKind::Normal | GroupKind::FieldResult | GroupKind::FieldInst => {
                    append_text_byte(ctx, byte);
                }
                GroupKind::Object | GroupKind::Picture => {
                    if byte.is_ascii_hexdigit() {
                        if let Some(obj) = ctx.object_stack.last_mut() {
                            obj.data_hex_len += 1;
                            if ctx.max_object_hex_len > 0
                                && obj.data_hex_len > ctx.max_object_hex_len
                            {
                                return Err(ParseError::ResourceLimit(format!(
                                    "RTF objdata too large: {} hex chars (max: {})",
                                    obj.data_hex_len, ctx.max_object_hex_len
                                )));
                            }
                        }
                    }
                }
                GroupKind::FontTable | GroupKind::ColorTable => {
                    append_text_byte(ctx, byte);
                }
                _ => {}
            },
            None => break,
        }
    }
    flush_text(ctx, store, None)?;
    finalize_table_if_open(ctx, store);
    finalize_paragraph(ctx, store);
    finalize_section(ctx, store);
    Ok(())
}

fn parse_control(
    cursor: &mut RtfCursor<'_>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<(), ParseError> {
    let Some(next) = cursor.next() else {
        return Ok(());
    };

    match next {
        b'\\' | b'{' | b'}' => {
            append_text_byte(ctx, next);
            return Ok(());
        }
        b'\'' => {
            let hi = cursor.next().unwrap_or(b'0');
            let lo = cursor.next().unwrap_or(b'0');
            if let (Some(h), Some(l)) = (hex_val(hi), hex_val(lo)) {
                append_text_byte(ctx, (h << 4) | l);
            }
            return Ok(());
        }
        b'*' => {
            if let Some(group) = ctx.group_stack.last_mut() {
                group.kind = GroupKind::Skip;
            }
            return Ok(());
        }
        b'~' => {
            append_text(ctx, " ");
            return Ok(());
        }
        b'-' => {
            return Ok(()); // optional hyphen
        }
        b'_' => {
            append_text(ctx, "-");
            return Ok(());
        }
        b'\n' | b'\r' => return Ok(()),
        _ => {}
    }

    if next.is_ascii_alphabetic() {
        let (word, param) = parse_control_word_and_param(cursor, next);
        handle_control_word(&word, param, ctx, store)?;
    }

    Ok(())
}

fn parse_control_word_and_param(cursor: &mut RtfCursor<'_>, first: u8) -> (String, Option<i32>) {
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
        if b.is_ascii_digit() {
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

fn handle_control_word(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<(), ParseError> {
    if handle_paragraph_controls(word, param, ctx, store)? {
        return Ok(());
    }
    if handle_run_style_controls(word, param, ctx) {
        return Ok(());
    }
    if handle_group_controls(word, param, ctx)? {
        return Ok(());
    }
    if handle_field_controls(word, ctx) {
        return Ok(());
    }
    if handle_object_controls(word, param, ctx) {
        return Ok(());
    }
    if handle_table_controls(word, param, ctx, store)? {
        return Ok(());
    }
    if handle_encoding_controls(word, param, ctx) {
        return Ok(());
    }
    Ok(())
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
            ctx.current_props.bold = Some(param.unwrap_or(1) != 0);
        }
        "i" => {
            ctx.current_props.italic = Some(param.unwrap_or(1) != 0);
        }
        "ul" => {
            ctx.current_props.underline = Some(param.unwrap_or(1) != 0);
        }
        "ulnone" => {
            ctx.current_props.underline = Some(false);
        }
        "strike" => {
            ctx.current_props.strike = Some(param.unwrap_or(1) != 0);
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

fn flush_text(
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
    store.insert(IRNode::Run(run));

    attach_flushed_run(ctx, store, &text, run_id);

    Ok(())
}

fn flush_stylesheet_text(ctx: &mut RtfParseContext, text: &str) {
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

fn flush_object_text(ctx: &mut RtfParseContext, text: &str) {
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

fn attach_flushed_run(ctx: &mut RtfParseContext, store: &mut IrStore, text: &str, run_id: NodeId) {
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

fn append_text(ctx: &mut RtfParseContext, text: &str) {
    ctx.current_text.push_str(text);
}

fn append_text_byte(ctx: &mut RtfParseContext, byte: u8) {
    let binding = [byte];
    let (text, _, _) = ctx.encoding.decode(&binding);
    ctx.current_text.push_str(&text);
}

fn ensure_paragraph(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if ctx.current_paragraph.is_none() {
        let mut para = Paragraph::new();
        para.span = Some(SourceSpan::new("rtf"));
        apply_pending_paragraph(&mut para, ctx);
        ctx.current_paragraph = Some(para);
    }
    ensure_section(ctx, store);
}

fn ensure_section(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if ctx.current_section.is_none() {
        let mut section = docir_core::ir::Section::new();
        section.span = Some(SourceSpan::new("rtf"));
        let section_id = section.id;
        store.insert(IRNode::Section(section));
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

fn pending_numbering(ctx: &RtfParseContext) -> Option<NumberingInfo> {
    let list_override = ctx.pending_list_override?;
    let level = ctx.pending_list_level.unwrap_or(0);
    let list_id = ctx
        .list_overrides
        .get(&list_override)
        .copied()
        .unwrap_or(list_override);
    let format = ctx.list_level_formats.get(&(list_id, level)).cloned();
    Some(NumberingInfo {
        num_id: list_id as u32,
        level,
        format,
    })
}

fn color_from_index(ctx: &RtfParseContext, index: usize) -> Option<String> {
    ctx.color_table.colors.get(index).and_then(|c| c.clone())
}

fn set_border_target(ctx: &mut RtfParseContext, target: BorderTarget) {
    ctx.pending_border_target = Some(target);
}

fn apply_border(ctx: &mut RtfParseContext) {
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

fn apply_paragraph_border(ctx: &mut RtfParseContext) {
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

fn cell_width_from_row(ctx: &RtfParseContext) -> Option<TableWidth> {
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

fn finalize_paragraph(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if let Some(para) = ctx.current_paragraph.take() {
        let para_id = para.id;
        store.insert(IRNode::Paragraph(para));
        if let Some(cell) = ctx.current_cell.as_mut() {
            cell.content.push(para_id);
        } else if let Some(section_id) = ctx.current_section {
            if let Some(IRNode::Section(section)) = store.get_mut(section_id) {
                section.content.push(para_id);
            }
        }
    }
    ctx.pending_list_override = None;
    ctx.pending_list_level = None;
}

fn finalize_section(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if let Some(section_id) = ctx.current_section.take() {
        if store.get(section_id).is_none() {
            let mut section = docir_core::ir::Section::new();
            section.id = section_id;
            section.span = Some(SourceSpan::new("rtf"));
            store.insert(IRNode::Section(section));
        }
    }
}

fn finalize_cell(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if let Some(mut cell) = ctx.current_cell.take() {
        if let Some(mut props) = ctx.pending_cell_props.take() {
            if let Some(width) = cell_width_from_row(ctx) {
                props.width = Some(width);
            }
            cell.properties = props;
        } else if let Some(width) = cell_width_from_row(ctx) {
            cell.properties.width = Some(width);
        }
        ctx.current_cell_index = ctx.current_cell_index.saturating_add(1);
        let cell_id = cell.id;
        store.insert(IRNode::TableCell(cell));
        if let Some(row) = ctx.current_row.as_mut() {
            row.cells.push(cell_id);
        }
    }
}

fn finalize_row(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if let Some(row) = ctx.current_row.take() {
        let row_id = row.id;
        store.insert(IRNode::TableRow(row));
        if let Some(table) = ctx.current_table.as_mut() {
            table.rows.push(row_id);
        }
    }
    ctx.row_cellx.clear();
    ctx.current_cell_index = 0;
}

fn finalize_table_if_open(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if let Some(table) = ctx.current_table.take() {
        let table_id = table.id;
        store.insert(IRNode::Table(table));
        ensure_section(ctx, store);
        if let Some(section_id) = ctx.current_section {
            if let Some(IRNode::Section(section)) = store.get_mut(section_id) {
                section.content.push(table_id);
            }
        }
    }
}

fn finalize_field(ctx: &mut RtfParseContext, store: &mut IrStore) {
    let Some(field) = ctx.field_stack.pop() else {
        return;
    };
    let instr = field.instruction.trim();
    let mut instruction = if instr.is_empty() {
        None
    } else {
        Some(instr.to_string())
    };
    if let Some(instr_text) = instruction.clone() {
        if let Some((target, _args, _switches)) = parse_hyperlink_instruction(&instr_text) {
            let mut link = Hyperlink::new(target, true);
            link.runs = field.runs.clone();
            let link_id = link.id;
            store.insert(IRNode::Hyperlink(link));
            ensure_paragraph(ctx, store);
            if let Some(para) = ctx.current_paragraph.as_mut() {
                para.runs.push(link_id);
            }
            if let Some(ext_id) = create_external_ref(&instr_text, store, ctx) {
                ctx.external_refs.push(ext_id);
            }
            return;
        }
    }

    let mut node = Field::new(instruction.take());
    node.runs = field.runs.clone();
    if let Some(instr_text) = node.instruction.as_ref() {
        node.instruction_parsed = parse_field_instruction(instr_text);
    }
    let field_id = node.id;
    store.insert(IRNode::Field(node));
    ensure_paragraph(ctx, store);
    if let Some(para) = ctx.current_paragraph.as_mut() {
        para.runs.push(field_id);
    }
}

fn finalize_style_entry(ctx: &mut RtfParseContext) {
    let Some(style_ctx) = ctx.current_style.take() else {
        return;
    };
    let name = style_ctx.name_buf.trim().to_string();
    if !name.is_empty() {
        push_style_from_ctx(ctx, &style_ctx, name);
    }
}

fn push_style_from_ctx(ctx: &mut RtfParseContext, style_ctx: &StyleEntryContext, name: String) {
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

fn finalize_list_entry(ctx: &mut RtfParseContext) {
    if let Some(list_id) = ctx.current_list_id {
        let levels = ctx.current_list_level.max(1);
        ctx.list_levels.insert(list_id, levels);
    }
    ctx.current_list_id = None;
    ctx.current_list_level = 0;
}

fn finalize_list_override(ctx: &mut RtfParseContext) {
    if let (Some(override_id), Some(list_id)) =
        (ctx.current_list_override, ctx.current_list_override_list_id)
    {
        ctx.list_overrides.insert(override_id, list_id);
    }
    ctx.current_list_override = None;
    ctx.current_list_override_list_id = None;
}

fn parse_field_instruction(text: &str) -> Option<FieldInstruction> {
    parse_field_instruction_impl(text)
}

fn parse_hyperlink_instruction(text: &str) -> Option<(String, Vec<String>, Vec<String>)> {
    parse_hyperlink_instruction_impl(text)
}

fn tokenize_field_instruction(text: &str) -> Vec<String> {
    tokenize_field_instruction_impl(text)
}

fn create_external_ref(
    instr: &str,
    store: &mut IrStore,
    _ctx: &mut RtfParseContext,
) -> Option<NodeId> {
    if let Some((target, _, _)) = parse_hyperlink_instruction(instr) {
        let mut ext = ExternalReference::new(ExternalRefType::Hyperlink, target.clone());
        ext.span = Some(SourceSpan::new("rtf"));
        let id = ext.id;
        store.insert(IRNode::ExternalReference(ext));
        return Some(id);
    }
    None
}

fn run_properties_from_state(ctx: &RtfParseContext) -> RunProperties {
    let mut props = RunProperties::default();
    props.style_id = ctx.current_props.style_id.clone();
    props.bold = ctx.current_props.bold;
    props.italic = ctx.current_props.italic;
    props.underline = ctx.current_props.underline.map(|u| {
        if u {
            docir_core::ir::UnderlineStyle::Single
        } else {
            docir_core::ir::UnderlineStyle::None
        }
    });
    props.strike = ctx.current_props.strike;
    props.font_size = ctx.current_props.font_size;
    props.vertical_align = ctx.current_props.vertical;
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
