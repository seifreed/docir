use super::{
    attr_any, note_kind_from_local, HwpxNoteKind, HwpxNoteState, HwpxSectionState, NodeId,
    Paragraph, SourceSpan,
};
use docir_core::ir::{Comment, CommentReference, Endnote, Footnote, IRNode};
use docir_core::visitor::IrStore;
use quick_xml::events::BytesStart;

pub(super) fn start_hwpx_note(
    e: &BytesStart<'_>,
    local: &[u8],
    state: &mut HwpxSectionState,
) -> bool {
    let Some(kind) = note_kind_from_local(local) else {
        return false;
    };
    let id = attr_any(
        e,
        &[b"id", b"commentId", b"comment-id", b"refId", b"ref-id"],
    )
    .unwrap_or_else(|| next_hwpx_note_id(kind, state));
    state.note_stack.push(HwpxNoteState {
        kind,
        id,
        author: attr_any(e, &[b"author", b"writer"]),
        date: attr_any(e, &[b"date", b"created", b"time"]),
        parent: attr_any(e, &[b"parent", b"parentId", b"parent-id"]),
        content: Vec::new(),
        current_para: None,
    });
    true
}

pub(super) fn push_hwpx_comment_reference(
    e: &BytesStart<'_>,
    local: &[u8],
    source: &str,
    store: &mut IrStore,
    state: &mut HwpxSectionState,
) -> bool {
    if !matches!(
        local,
        b"commentRef" | b"comment-ref" | b"annotationRef" | b"noteRef" | b"note-ref"
    ) {
        return false;
    }
    if let Some(comment_id) = attr_any(e, &[b"id", b"ref", b"refId", b"ref-id"]) {
        let mut node = CommentReference::new(comment_id);
        node.span = Some(SourceSpan::new(source));
        let node_id = node.id;
        store.insert(IRNode::CommentReference(node));
        push_node_to_hwpx_context(
            node_id,
            &mut state.current_para,
            state.note_stack.as_mut_slice(),
            source,
        );
    }
    true
}

pub(super) fn close_hwpx_note(
    local: &[u8],
    source: &str,
    store: &mut IrStore,
    comments: &mut Vec<NodeId>,
    footnotes: &mut Vec<NodeId>,
    endnotes: &mut Vec<NodeId>,
    state: &mut HwpxSectionState,
) -> bool {
    let Some(kind) = note_kind_from_local(local) else {
        return false;
    };
    let Some(mut note) = state.note_stack.pop() else {
        return true;
    };
    finalize_note_paragraph(&mut note, store);
    emit_note_node(kind, note, source, store, comments, footnotes, endnotes);
    true
}

pub(super) fn flush_hwpx_notes(
    source: &str,
    store: &mut IrStore,
    comments: &mut Vec<NodeId>,
    footnotes: &mut Vec<NodeId>,
    endnotes: &mut Vec<NodeId>,
    state: &mut HwpxSectionState,
) {
    while let Some(mut note) = state.note_stack.pop() {
        finalize_note_paragraph(&mut note, store);
        emit_note_node(
            note.kind, note, source, store, comments, footnotes, endnotes,
        );
    }
}

pub(super) fn finalize_note_paragraph(note: &mut HwpxNoteState, store: &mut IrStore) {
    if let Some(para) = note.current_para.take() {
        let para_id = para.id;
        store.insert(IRNode::Paragraph(para));
        note.content.push(para_id);
    }
}

pub(super) fn push_node_to_hwpx_context(
    node_id: NodeId,
    current_para: &mut Option<Paragraph>,
    note_stack: &mut [HwpxNoteState],
    source: &str,
) {
    if let Some(note) = note_stack.last_mut() {
        if note.current_para.is_none() {
            let mut para = Paragraph::new();
            para.span = Some(SourceSpan::new(source));
            note.current_para = Some(para);
        }
        if let Some(para) = note.current_para.as_mut() {
            para.runs.push(node_id);
        }
    } else {
        if current_para.is_none() {
            let mut para = Paragraph::new();
            para.span = Some(SourceSpan::new(source));
            *current_para = Some(para);
        }
        if let Some(para) = current_para.as_mut() {
            para.runs.push(node_id);
        }
    }
}

fn next_hwpx_note_id(kind: HwpxNoteKind, state: &mut HwpxSectionState) -> String {
    match kind {
        HwpxNoteKind::Comment => {
            state.comment_counter = state.comment_counter.saturating_add(1);
            format!("hwpx-comment-{}", state.comment_counter)
        }
        HwpxNoteKind::Footnote => {
            state.footnote_counter = state.footnote_counter.saturating_add(1);
            format!("hwpx-footnote-{}", state.footnote_counter)
        }
        HwpxNoteKind::Endnote => {
            state.endnote_counter = state.endnote_counter.saturating_add(1);
            format!("hwpx-endnote-{}", state.endnote_counter)
        }
    }
}

fn emit_note_node(
    kind: HwpxNoteKind,
    note: HwpxNoteState,
    source: &str,
    store: &mut IrStore,
    comments: &mut Vec<NodeId>,
    footnotes: &mut Vec<NodeId>,
    endnotes: &mut Vec<NodeId>,
) {
    match kind {
        HwpxNoteKind::Comment => {
            let mut comment = Comment::new(note.id);
            comment.author = note.author;
            comment.date = note.date;
            comment.parent_id = note.parent;
            comment.content = note.content;
            comment.span = Some(SourceSpan::new(source));
            let comment_id = comment.id;
            store.insert(IRNode::Comment(comment));
            comments.push(comment_id);
        }
        HwpxNoteKind::Footnote => {
            let mut footnote = Footnote::new(note.id);
            footnote.content = note.content;
            footnote.span = Some(SourceSpan::new(source));
            let footnote_id = footnote.id;
            store.insert(IRNode::Footnote(footnote));
            footnotes.push(footnote_id);
        }
        HwpxNoteKind::Endnote => {
            let mut endnote = Endnote::new(note.id);
            endnote.content = note.content;
            endnote.span = Some(SourceSpan::new(source));
            let endnote_id = endnote.id;
            store.insert(IRNode::Endnote(endnote));
            endnotes.push(endnote_id);
        }
    }
}
