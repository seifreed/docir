use crate::error::ParseError;
#[path = "section_parse_events.rs"]
mod section_parse_events;
#[path = "section_parse_notes.rs"]
mod section_parse_notes;
#[path = "section_parse_revisions.rs"]
mod section_parse_revisions;
use crate::hwp::{
    attr_any, finalize_cell_hwpx, finalize_row_hwpx, finalize_table_hwpx, parse_hwpx_shape,
    run_properties_from_attrs,
};
use crate::xml_utils::{local_name, xml_error};
use docir_core::ir::{
    IRNode, NumberingInfo, Paragraph, Revision, RevisionType, Run, RunProperties, Table, TableCell,
    TableRow,
};
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use section_parse_notes::{
    close_hwpx_note, finalize_note_paragraph, flush_hwpx_notes, push_hwpx_comment_reference,
    push_node_to_hwpx_context, start_hwpx_note,
};
use section_parse_revisions::{close_hwpx_revision, flush_hwpx_revisions, push_hwpx_revision};
use std::collections::HashMap;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum HwpxNoteKind {
    Comment,
    Footnote,
    Endnote,
}

struct HwpxNoteState {
    kind: HwpxNoteKind,
    id: String,
    author: Option<String>,
    date: Option<String>,
    parent: Option<String>,
    content: Vec<NodeId>,
    current_para: Option<Paragraph>,
}

struct HwpxSectionState {
    content: Vec<NodeId>,
    current_para: Option<Paragraph>,
    current_table: Option<Table>,
    current_row: Option<TableRow>,
    current_cell: Option<TableCell>,
    in_text: bool,
    run_props: Option<RunProperties>,
    list_level: u32,
    revision_stack: Vec<Revision>,
    note_stack: Vec<HwpxNoteState>,
    comment_counter: u32,
    footnote_counter: u32,
    endnote_counter: u32,
}

impl HwpxSectionState {
    fn new() -> Self {
        Self {
            content: Vec::new(),
            current_para: None,
            current_table: None,
            current_row: None,
            current_cell: None,
            in_text: false,
            run_props: None,
            list_level: 0,
            revision_stack: Vec::new(),
            note_stack: Vec::new(),
            comment_counter: 0,
            footnote_counter: 0,
            endnote_counter: 0,
        }
    }
}

pub(crate) fn note_kind_from_local(local: &[u8]) -> Option<HwpxNoteKind> {
    match local {
        b"comment" | b"annotation" | b"note" => Some(HwpxNoteKind::Comment),
        b"footnote" => Some(HwpxNoteKind::Footnote),
        b"endnote" => Some(HwpxNoteKind::Endnote),
        _ => None,
    }
}

pub(crate) fn revision_type_from_local(local: &[u8]) -> Option<RevisionType> {
    match local {
        b"ins" | b"insert" => Some(RevisionType::Insert),
        b"del" | b"delete" => Some(RevisionType::Delete),
        b"moveFrom" | b"move-from" => Some(RevisionType::MoveFrom),
        b"moveTo" | b"move-to" => Some(RevisionType::MoveTo),
        b"formatChange" | b"format-change" => Some(RevisionType::FormatChange),
        _ => None,
    }
}

pub(crate) fn parse_hwpx_section(
    xml: &str,
    source: &str,
    store: &mut IrStore,
    comments: &mut Vec<NodeId>,
    footnotes: &mut Vec<NodeId>,
    endnotes: &mut Vec<NodeId>,
    media_lookup: &HashMap<String, NodeId>,
) -> Result<Vec<NodeId>, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut state = HwpxSectionState::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                handle_hwpx_start(&e, source, store, media_lookup, &mut state)?;
            }
            Ok(Event::Empty(e)) => {
                handle_hwpx_empty(&e, source, store, media_lookup, &mut state)?;
            }
            Ok(Event::End(e)) => {
                handle_hwpx_end(&e, source, store, comments, footnotes, endnotes, &mut state);
            }
            Ok(Event::Text(e)) => {
                handle_hwpx_text(&e, source, store, &mut state);
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(source, e));
            }
            _ => {}
        }
        buf.clear();
    }

    flush_hwpx_notes(source, store, comments, footnotes, endnotes, &mut state);
    flush_hwpx_revisions(source, store, &mut state);
    finalize_paragraph_hwpx(
        &mut state.current_para,
        &mut state.current_cell,
        &mut state.content,
        store,
    );
    finalize_table_hwpx(
        &mut state.current_table,
        &mut state.current_cell,
        &mut state.content,
        store,
    );

    Ok(state.content)
}

fn handle_hwpx_start(
    e: &BytesStart<'_>,
    source: &str,
    store: &mut IrStore,
    media_lookup: &HashMap<String, NodeId>,
    state: &mut HwpxSectionState,
) -> Result<(), ParseError> {
    section_parse_events::handle_hwpx_start(e, source, store, media_lookup, state)
}

fn push_hwpx_shape(
    e: &BytesStart<'_>,
    local: &[u8],
    source: &str,
    store: &mut IrStore,
    media_lookup: &HashMap<String, NodeId>,
    state: &mut HwpxSectionState,
) -> bool {
    let Some(shape_id) = parse_hwpx_shape(e, local, source, media_lookup, store) else {
        return false;
    };
    push_node_to_hwpx_context(
        shape_id,
        &mut state.current_para,
        state.note_stack.as_mut_slice(),
        source,
    );
    true
}

fn handle_hwpx_empty(
    e: &BytesStart<'_>,
    source: &str,
    store: &mut IrStore,
    media_lookup: &HashMap<String, NodeId>,
    state: &mut HwpxSectionState,
) -> Result<(), ParseError> {
    section_parse_events::handle_hwpx_empty(e, source, store, media_lookup, state)
}

fn handle_hwpx_end(
    e: &quick_xml::events::BytesEnd<'_>,
    source: &str,
    store: &mut IrStore,
    comments: &mut Vec<NodeId>,
    footnotes: &mut Vec<NodeId>,
    endnotes: &mut Vec<NodeId>,
    state: &mut HwpxSectionState,
) {
    section_parse_events::handle_hwpx_end(e, source, store, comments, footnotes, endnotes, state)
}

fn handle_hwpx_text(
    e: &quick_xml::events::BytesText<'_>,
    source: &str,
    store: &mut IrStore,
    state: &mut HwpxSectionState,
) {
    section_parse_events::handle_hwpx_text(e, source, store, state)
}

pub(super) fn finalize_paragraph_hwpx(
    current_para: &mut Option<Paragraph>,
    current_cell: &mut Option<TableCell>,
    content: &mut Vec<NodeId>,
    store: &mut IrStore,
) {
    if let Some(para) = current_para.take() {
        let para_id = para.id;
        store.insert(IRNode::Paragraph(para));
        if let Some(cell) = current_cell.as_mut() {
            cell.content.push(para_id);
        } else {
            content.push(para_id);
        }
    }
}
