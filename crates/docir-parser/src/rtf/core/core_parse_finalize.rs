#[path = "core_parse_finalize_field.rs"]
mod core_parse_finalize_field;
#[path = "core_parse_finalize_group.rs"]
mod core_parse_finalize_group;
use super::super::field_utils::{
    parse_field_instruction as parse_field_instruction_impl,
    parse_hyperlink_instruction as parse_hyperlink_instruction_impl,
};
use super::super::state::{GroupKind, StyleEntryContext};
use super::{cell_width_from_row, ensure_paragraph, ensure_section, IrStore, RtfParseContext};
use crate::rtf::objects::{finalize_object, finalize_picture};
use docir_core::ir::{FieldInstruction, IRNode};
use docir_core::types::SourceSpan;

pub(crate) fn finalize_paragraph(ctx: &mut RtfParseContext, store: &mut IrStore) {
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

pub(crate) fn finalize_section(ctx: &mut RtfParseContext, store: &mut IrStore) {
    if let Some(section_id) = ctx.current_section.take() {
        if store.get(section_id).is_none() {
            let mut section = docir_core::ir::Section::new();
            section.id = section_id;
            section.span = Some(SourceSpan::new("rtf"));
            store.insert(IRNode::Section(section));
        }
    }
}

pub(crate) fn finalize_cell(ctx: &mut RtfParseContext, store: &mut IrStore) {
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

pub(crate) fn finalize_row(ctx: &mut RtfParseContext, store: &mut IrStore) {
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

pub(crate) fn finalize_table_if_open(ctx: &mut RtfParseContext, store: &mut IrStore) {
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

pub(crate) fn finalize_field(ctx: &mut RtfParseContext, store: &mut IrStore) {
    core_parse_finalize_field::finalize_field(ctx, store)
}

pub(crate) fn push_style_from_ctx(
    ctx: &mut RtfParseContext,
    style_ctx: &StyleEntryContext,
    name: String,
) {
    core_parse_finalize_group::push_style_from_ctx(ctx, style_ctx, name)
}

#[cfg(test)]
pub(crate) fn finalize_list_entry(ctx: &mut RtfParseContext) {
    core_parse_finalize_group::finalize_list_entry(ctx)
}

#[cfg(test)]
pub(crate) fn finalize_list_override(ctx: &mut RtfParseContext) {
    core_parse_finalize_group::finalize_list_override(ctx)
}

pub(crate) fn parse_field_instruction(text: &str) -> Option<FieldInstruction> {
    parse_field_instruction_impl(text)
}

pub(crate) fn handle_group_end(ctx: &mut RtfParseContext, store: &mut IrStore) {
    core_parse_finalize_group::handle_group_end(ctx, store)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::rtf::core::state::{GroupKind, StyleEntryContext};
    use docir_core::ir::{IRNode, StyleType};

    #[test]
    fn finalize_list_entry_and_override_persist_values() {
        let mut ctx = RtfParseContext::new(128, 1024 * 1024);
        let mut store = IrStore::new();

        ctx.current_list_id = Some(7);
        ctx.current_list_level = 3;
        finalize_list_entry(&mut ctx);
        assert_eq!(ctx.list_levels.get(&7), Some(&3));

        ctx.current_list_override = Some(9);
        ctx.current_list_override_list_id = Some(7);
        finalize_list_override(&mut ctx);
        assert_eq!(ctx.list_overrides.get(&9), Some(&7));

        handle_group_end(&mut ctx, &mut store);
    }

    #[test]
    fn handle_group_end_stylesheet_entry_pushes_style() {
        let mut ctx = RtfParseContext::new(128, 1024 * 1024);
        let mut store = IrStore::new();

        ctx.current_style = Some(StyleEntryContext {
            style_id: Some("Heading1".to_string()),
            style_type: Some(StyleType::Paragraph),
            name_buf: "Heading 1".to_string(),
        });
        ctx.push_group(GroupKind::StylesheetEntry)
            .expect("group push should succeed");
        handle_group_end(&mut ctx, &mut store);

        let set = ctx.style_set.expect("style set should be created");
        assert_eq!(set.styles.len(), 1);
        assert_eq!(set.styles[0].style_id, "Heading1");
        assert_eq!(set.styles[0].name.as_deref(), Some("Heading 1"));
    }

    #[test]
    fn finalize_field_builds_hyperlink_and_external_reference() {
        let mut ctx = RtfParseContext::new(128, 1024 * 1024);
        let mut store = IrStore::new();

        ctx.field_stack.push(crate::rtf::core::state::FieldContext {
            instruction: r#"HYPERLINK "https://example.test""#.to_string(),
            runs: Vec::new(),
        });
        finalize_field(&mut ctx, &mut store);

        assert_eq!(ctx.external_refs.len(), 1);
        let mut hyperlink_found = false;
        for (_, node) in store.iter() {
            if matches!(node, IRNode::Hyperlink(_)) {
                hyperlink_found = true;
                break;
            }
        }
        assert!(hyperlink_found, "expected hyperlink node in store");
    }
}
