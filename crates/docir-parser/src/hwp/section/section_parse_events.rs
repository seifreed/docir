use super::{
    attr_any, close_hwpx_note, close_hwpx_revision, finalize_cell_hwpx, finalize_note_paragraph,
    finalize_paragraph_hwpx, finalize_row_hwpx, finalize_table_hwpx, local_name,
    push_hwpx_comment_reference, push_hwpx_revision, push_hwpx_shape, push_node_to_hwpx_context,
    run_properties_from_attrs, HwpxSectionState, IRNode, IrStore, NodeId, NumberingInfo, Paragraph,
    ParseError, Run, SourceSpan, Table, TableCell, TableRow,
};
use quick_xml::events::{BytesEnd, BytesStart, BytesText};
use std::collections::HashMap;

pub(super) fn handle_hwpx_start(
    e: &BytesStart<'_>,
    source: &str,
    store: &mut IrStore,
    media_lookup: &HashMap<String, NodeId>,
    state: &mut HwpxSectionState,
) -> Result<(), ParseError> {
    let name = e.name().as_ref().to_vec();
    let local = local_name(&name);
    if super::start_hwpx_note(e, local, state) {
        return Ok(());
    }
    if push_hwpx_comment_reference(e, local, source, store, state) {
        return Ok(());
    }
    if push_hwpx_revision(e, local, source, state) {
        return Ok(());
    }
    if push_hwpx_shape(e, local, source, store, media_lookup, state) {
        return Ok(());
    }
    apply_hwpx_start_element(local, e, source, store, state);
    Ok(())
}

pub(super) fn apply_hwpx_start_element(
    local: &[u8],
    e: &BytesStart<'_>,
    source: &str,
    store: &mut IrStore,
    state: &mut HwpxSectionState,
) {
    match local {
        b"p" => start_hwpx_paragraph(e, source, store, state),
        b"t" => state.in_text = true,
        b"r" | b"run" | b"span" => state.run_props = Some(run_properties_from_attrs(e)),
        b"tbl" | b"table" => {
            finalize_paragraph_hwpx(
                &mut state.current_para,
                &mut state.current_cell,
                &mut state.content,
                store,
            );
            if state.current_table.is_none() {
                state.current_table = Some(Table::new());
            }
        }
        b"tr" | b"row" if state.current_table.is_some() => {
            state.current_row = Some(TableRow::new());
        }
        b"tc" | b"cell" if state.current_row.is_some() => {
            state.current_cell = Some(TableCell::new());
        }
        b"list" | b"ul" | b"ol" => state.list_level = state.list_level.saturating_add(1),
        b"li" | b"list-item" => ensure_list_item_paragraph(source, state),
        _ => {}
    }
}

pub(super) fn start_hwpx_paragraph(
    e: &BytesStart<'_>,
    source: &str,
    store: &mut IrStore,
    state: &mut HwpxSectionState,
) {
    if let Some(note) = state.note_stack.last_mut() {
        finalize_note_paragraph(note, store);
        let mut para = Paragraph::new();
        para.span = Some(SourceSpan::new(source));
        note.current_para = Some(para);
        return;
    }
    finalize_paragraph_hwpx(
        &mut state.current_para,
        &mut state.current_cell,
        &mut state.content,
        store,
    );
    let mut para = Paragraph::new();
    para.span = Some(SourceSpan::new(source));
    if let Some(style_id) = attr_any(e, &[b"styleId", b"style-id", b"style"]) {
        para.style_id = Some(style_id);
    }
    apply_hwpx_numbering(&mut para, state.list_level);
    state.current_para = Some(para);
}

pub(super) fn ensure_list_item_paragraph(source: &str, state: &mut HwpxSectionState) {
    if !state.note_stack.is_empty() || state.current_para.is_some() {
        return;
    }
    let mut para = Paragraph::new();
    para.span = Some(SourceSpan::new(source));
    apply_hwpx_numbering(&mut para, state.list_level);
    state.current_para = Some(para);
}

pub(super) fn apply_hwpx_numbering(paragraph: &mut Paragraph, list_level: u32) {
    if list_level == 0 {
        return;
    }
    paragraph.properties.numbering = Some(NumberingInfo {
        num_id: 1,
        level: list_level - 1,
        format: None,
    });
}

pub(super) fn handle_hwpx_empty(
    e: &BytesStart<'_>,
    source: &str,
    store: &mut IrStore,
    media_lookup: &HashMap<String, NodeId>,
    state: &mut HwpxSectionState,
) -> Result<(), ParseError> {
    let name = e.name().as_ref().to_vec();
    let local = local_name(&name);
    if push_hwpx_comment_reference(e, local, source, store, state) {
        return Ok(());
    }
    let _ = push_hwpx_shape(e, local, source, store, media_lookup, state);
    Ok(())
}

pub(super) fn handle_hwpx_end(
    e: &BytesEnd<'_>,
    source: &str,
    store: &mut IrStore,
    comments: &mut Vec<NodeId>,
    footnotes: &mut Vec<NodeId>,
    endnotes: &mut Vec<NodeId>,
    state: &mut HwpxSectionState,
) {
    let name = e.name().as_ref().to_vec();
    let local = local_name(&name);
    if close_hwpx_note(local, source, store, comments, footnotes, endnotes, state) {
        return;
    }
    if close_hwpx_revision(local, source, store, state) {
        return;
    }
    apply_hwpx_end_element(local, store, state);
}

pub(super) fn apply_hwpx_end_element(
    local: &[u8],
    store: &mut IrStore,
    state: &mut HwpxSectionState,
) {
    match local {
        b"p" => {
            if let Some(note) = state.note_stack.last_mut() {
                finalize_note_paragraph(note, store);
            } else {
                finalize_paragraph_hwpx(
                    &mut state.current_para,
                    &mut state.current_cell,
                    &mut state.content,
                    store,
                );
            }
        }
        b"t" => state.in_text = false,
        b"r" | b"run" | b"span" => state.run_props = None,
        b"tc" | b"cell" => {
            finalize_cell_hwpx(&mut state.current_cell, &mut state.current_row, store)
        }
        b"tr" | b"row" => {
            finalize_row_hwpx(&mut state.current_row, &mut state.current_table, store)
        }
        b"tbl" | b"table" => finalize_table_hwpx(
            &mut state.current_table,
            &mut state.current_cell,
            &mut state.content,
            store,
        ),
        b"list" | b"ul" | b"ol" => state.list_level = state.list_level.saturating_sub(1),
        _ => {}
    }
}

pub(super) fn handle_hwpx_text(
    e: &BytesText<'_>,
    source: &str,
    store: &mut IrStore,
    state: &mut HwpxSectionState,
) {
    if state.in_text {
        let text = e.unescape().unwrap_or_default().to_string();
        if !text.is_empty() {
            let props = state.run_props.clone().unwrap_or_default();
            let mut run = Run::with_properties(text, props);
            run.span = Some(SourceSpan::new(source));
            let run_id = run.id;
            store.insert(IRNode::Run(run));
            if let Some(revision) = state.revision_stack.last_mut() {
                revision.content.push(run_id);
            } else {
                push_node_to_hwpx_context(
                    run_id,
                    &mut state.current_para,
                    state.note_stack.as_mut_slice(),
                    source,
                );
            }
        }
    }
}
