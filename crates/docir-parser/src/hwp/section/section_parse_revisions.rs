use super::{
    attr_any, push_node_to_hwpx_context, revision_type_from_local, HwpxSectionState, SourceSpan,
};
use docir_core::ir::{IRNode, Revision};
use docir_core::visitor::IrStore;
use quick_xml::events::BytesStart;

pub(super) fn push_hwpx_revision(
    e: &BytesStart<'_>,
    local: &[u8],
    source: &str,
    state: &mut HwpxSectionState,
) -> bool {
    let Some(change_type) = revision_type_from_local(local) else {
        return false;
    };
    let mut revision = Revision::new(change_type);
    revision.revision_id = attr_any(e, &[b"id", b"revId", b"revisionId", b"revision-id"]);
    revision.author = attr_any(e, &[b"author", b"writer"]);
    revision.date = attr_any(e, &[b"date", b"created", b"time"]);
    revision.span = Some(SourceSpan::new(source));
    state.revision_stack.push(revision);
    true
}

pub(super) fn close_hwpx_revision(
    local: &[u8],
    source: &str,
    store: &mut IrStore,
    state: &mut HwpxSectionState,
) -> bool {
    if revision_type_from_local(local).is_none() {
        return false;
    }
    if let Some(revision) = state.revision_stack.pop() {
        let revision_id = revision.id;
        store.insert(IRNode::Revision(revision));
        push_node_to_hwpx_context(
            revision_id,
            &mut state.current_para,
            state.note_stack.as_mut_slice(),
            source,
        );
    }
    true
}

pub(super) fn flush_hwpx_revisions(
    source: &str,
    store: &mut IrStore,
    state: &mut HwpxSectionState,
) {
    while let Some(revision) = state.revision_stack.pop() {
        let revision_id = revision.id;
        store.insert(IRNode::Revision(revision));
        push_node_to_hwpx_context(
            revision_id,
            &mut state.current_para,
            state.note_stack.as_mut_slice(),
            source,
        );
    }
}
