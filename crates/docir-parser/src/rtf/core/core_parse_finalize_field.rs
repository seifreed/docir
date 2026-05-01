use super::{
    ensure_paragraph, parse_field_instruction, parse_hyperlink_instruction_impl, IRNode, IrStore,
    RtfParseContext, SourceSpan,
};
use docir_core::ir::{Field, Hyperlink};
use docir_core::security::{ExternalRefType, ExternalReference};
use docir_core::types::NodeId;

pub(super) fn finalize_field(ctx: &mut RtfParseContext, store: &mut IrStore) {
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
            if let Some(ext_id) = create_external_ref(&instr_text, store) {
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

pub(super) fn parse_hyperlink_instruction(
    text: &str,
) -> Option<(String, Vec<String>, Vec<String>)> {
    parse_hyperlink_instruction_impl(text)
}

pub(super) fn create_external_ref(instr: &str, store: &mut IrStore) -> Option<NodeId> {
    if let Some((target, _, _)) = parse_hyperlink_instruction(instr) {
        let mut ext = ExternalReference::new(ExternalRefType::Hyperlink, target.clone());
        ext.span = Some(SourceSpan::new("rtf"));
        let id = ext.id;
        store.insert(IRNode::ExternalReference(ext));
        return Some(id);
    }
    None
}
