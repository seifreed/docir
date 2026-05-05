//! ODF text parsing helpers.

use super::helpers::{
    parse_annotation, parse_draw_frame, parse_note, parse_table, parse_tracked_changes, ListContext,
};
use super::spreadsheet::parse_content_spreadsheet_fast;
use super::{
    parse_paragraph, BookmarkEnd, BookmarkStart, CommentReference, Field, FieldInstruction,
    FieldKind, IRNode, IrStore, NodeId, NumberingInfo, OdfContentResult, OdfLimitCounter,
    Paragraph, ParagraphProperties, ParseError, Run, Section,
};
use crate::xml_utils::{attr_value_by_suffix, local_name, xml_error};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::HashMap;

struct OdfTextState {
    in_text: bool,
    list_stack: Vec<ListContext>,
    list_id_map: HashMap<String, u32>,
    next_list_id: u32,
    comment_counter: u32,
    content_result: OdfContentResult,
    pending_inline_nodes: Vec<NodeId>,
    revisions: Vec<NodeId>,
}

impl OdfTextState {
    fn new() -> Self {
        Self {
            in_text: false,
            list_stack: Vec::new(),
            list_id_map: HashMap::new(),
            next_list_id: 1,
            comment_counter: 1,
            content_result: OdfContentResult::default(),
            pending_inline_nodes: Vec::new(),
            revisions: Vec::new(),
        }
    }
}

pub(super) fn parse_content_text(
    xml: &[u8],
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<OdfContentResult, ParseError> {
    if limits.fast_mode() {
        return parse_content_spreadsheet_fast(xml, store, limits);
    }
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut section = Section::new();
    section.name = Some("body".to_string());
    let mut state = OdfTextState::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                handle_text_start(&e, &mut reader, store, limits, &mut section, &mut state)?
            }
            Ok(Event::Empty(e)) => handle_text_empty(&e, store, &mut section, &mut state),
            Ok(Event::End(e)) => handle_text_end(&e, &mut state),
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("content.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    section.content.extend(state.revisions);
    let section_id = section.id;
    store.insert(IRNode::Section(section));
    let mut result = OdfContentResult::default();
    result.content.push(section_id);
    result.comments = state.content_result.comments;
    result.footnotes = state.content_result.footnotes;
    result.endnotes = state.content_result.endnotes;
    Ok(result)
}

fn handle_text_start(
    e: &BytesStart<'_>,
    reader: &mut Reader<std::io::Cursor<&[u8]>>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    section: &mut Section,
    state: &mut OdfTextState,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"text" => state.in_text = true,
        b"list" if state.in_text => push_text_list(state, e),
        b"p" | b"h" if state.in_text => {
            parse_text_paragraph(e, reader, store, limits, section, state)?
        }
        b"table" if state.in_text => parse_text_table(reader, store, limits, section, state)?,
        b"annotation" if state.in_text => {
            parse_text_annotation(reader, store, limits, section, state)?
        }
        b"note" if state.in_text => parse_text_note(e, reader, store, limits, state)?,
        b"bookmark-start" | b"reference-mark-start" if state.in_text => {
            push_bookmark_start(e, store, section);
        }
        b"bookmark-end" | b"reference-mark-end" if state.in_text => {
            push_bookmark_end(e, store, section);
        }
        b"date" if state.in_text => push_date_field(store, section),
        b"time" if state.in_text => push_time_field(store, section),
        b"frame" if state.in_text => push_draw_frame(e, reader, store, section)?,
        b"tracked-changes" if state.in_text => {
            parse_text_tracked_changes(reader, store, limits, state)?
        }
        _ => {}
    }
    Ok(())
}

fn push_text_list(state: &mut OdfTextState, e: &BytesStart<'_>) {
    let style_name = attr_value_by_suffix(e, &[b":style-name"]).unwrap_or_default();
    let num_id = state.list_id_map.entry(style_name).or_insert_with(|| {
        let id = state.next_list_id;
        state.next_list_id += 1;
        id
    });
    let level = state.list_stack.len() as u32;
    state.list_stack.push(ListContext {
        num_id: *num_id,
        level,
    });
}

fn parse_text_paragraph(
    e: &BytesStart<'_>,
    reader: &mut Reader<std::io::Cursor<&[u8]>>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    section: &mut Section,
    state: &mut OdfTextState,
) -> Result<(), ParseError> {
    let outline_level =
        attr_value_by_suffix(e, &[b":outline-level"]).and_then(|v| v.parse::<u8>().ok());
    let numbering = state.list_stack.last().map(|ctx| NumberingInfo {
        num_id: ctx.num_id,
        level: ctx.level,
        format: None,
    });
    let paragraph_id = parse_paragraph(
        reader,
        e.name().as_ref(),
        numbering,
        outline_level,
        store,
        &mut state.pending_inline_nodes,
        limits,
    )?;
    section.content.append(&mut state.pending_inline_nodes);
    section.content.push(paragraph_id);
    Ok(())
}

fn parse_text_table(
    reader: &mut Reader<std::io::Cursor<&[u8]>>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    section: &mut Section,
    state: &mut OdfTextState,
) -> Result<(), ParseError> {
    let table_id = parse_table(reader, store, limits)?;
    section.content.append(&mut state.pending_inline_nodes);
    section.content.push(table_id);
    Ok(())
}

fn parse_text_annotation(
    reader: &mut Reader<std::io::Cursor<&[u8]>>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    section: &mut Section,
    state: &mut OdfTextState,
) -> Result<(), ParseError> {
    let comment_id = format!("odf-annotation-{}", state.comment_counter);
    state.comment_counter += 1;
    let comment_node = parse_annotation(reader, &comment_id, store, limits)?;
    state.content_result.comments.push(comment_node);
    let comment_ref = CommentReference::new(comment_id);
    let ref_id = comment_ref.id;
    store.insert(IRNode::CommentReference(comment_ref));
    section.content.push(ref_id);
    Ok(())
}

fn parse_text_note(
    e: &BytesStart<'_>,
    reader: &mut Reader<std::io::Cursor<&[u8]>>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    state: &mut OdfTextState,
) -> Result<(), ParseError> {
    let note_class =
        attr_value_by_suffix(e, &[b":note-class"]).unwrap_or_else(|| "footnote".to_string());
    let note_id = format!("odf-note-{}", state.comment_counter);
    state.comment_counter += 1;
    let note = parse_note(reader, &note_id, &note_class, store, limits)?;
    match note_class.as_str() {
        "endnote" => state.content_result.endnotes.push(note),
        _ => state.content_result.footnotes.push(note),
    }
    Ok(())
}

fn push_draw_frame(
    e: &BytesStart<'_>,
    reader: &mut Reader<std::io::Cursor<&[u8]>>,
    store: &mut IrStore,
    section: &mut Section,
) -> Result<(), ParseError> {
    if let Some(shape_id) = parse_draw_frame(reader, e, store)? {
        section.content.push(shape_id);
    }
    Ok(())
}

fn parse_text_tracked_changes(
    reader: &mut Reader<std::io::Cursor<&[u8]>>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    state: &mut OdfTextState,
) -> Result<(), ParseError> {
    let mut tracked = parse_tracked_changes(reader, store, limits)?;
    state.revisions.append(&mut tracked);
    Ok(())
}

fn handle_text_empty(
    e: &BytesStart<'_>,
    store: &mut IrStore,
    section: &mut Section,
    state: &mut OdfTextState,
) {
    match local_name(e.name().as_ref()) {
        b"bookmark-start" if state.in_text => {
            push_bookmark_start(e, store, section);
        }
        b"bookmark-end" if state.in_text => {
            push_bookmark_end(e, store, section);
        }
        b"reference-mark-start" if state.in_text => {
            push_bookmark_start(e, store, section);
        }
        b"reference-mark-end" if state.in_text => {
            push_bookmark_end(e, store, section);
        }
        b"date" if state.in_text => {
            push_date_field(store, section);
        }
        b"time" if state.in_text => {
            push_time_field(store, section);
        }
        b"frame" if state.in_text => {}
        _ => {}
    }
}

fn handle_text_end(e: &quick_xml::events::BytesEnd<'_>, state: &mut OdfTextState) {
    match local_name(e.name().as_ref()) {
        b"text" => state.in_text = false,
        b"list" => {
            state.list_stack.pop();
        }
        _ => {}
    }
}

fn push_bookmark_start(e: &BytesStart<'_>, store: &mut IrStore, section: &mut Section) {
    if let Some(name) = attr_value_by_suffix(e, &[b":name"]) {
        let mut bookmark = BookmarkStart::new(name.clone());
        bookmark.name = Some(name);
        let bookmark_id = bookmark.id;
        store.insert(IRNode::BookmarkStart(bookmark));
        section.content.push(bookmark_id);
    }
}

fn push_bookmark_end(e: &BytesStart<'_>, store: &mut IrStore, section: &mut Section) {
    if let Some(name) = attr_value_by_suffix(e, &[b":name"]) {
        let bookmark = BookmarkEnd::new(name);
        let bookmark_id = bookmark.id;
        store.insert(IRNode::BookmarkEnd(bookmark));
        section.content.push(bookmark_id);
    }
}

fn push_date_field(store: &mut IrStore, section: &mut Section) {
    let mut field = Field::new(Some("DATE".to_string()));
    field.instruction_parsed = Some(FieldInstruction {
        kind: FieldKind::Date,
        args: Vec::new(),
        switches: Vec::new(),
    });
    let field_id = field.id;
    store.insert(IRNode::Field(field));
    section.content.push(field_id);
}

fn push_time_field(store: &mut IrStore, section: &mut Section) {
    let field = Field::new(Some("TIME".to_string()));
    let field_id = field.id;
    store.insert(IRNode::Field(field));
    section.content.push(field_id);
}

pub(super) fn build_paragraph(
    store: &mut IrStore,
    text: &str,
    numbering: Option<NumberingInfo>,
    outline_level: Option<u8>,
) -> NodeId {
    let mut paragraph = Paragraph::new();
    if numbering.is_some() || outline_level.is_some() {
        paragraph.properties = ParagraphProperties {
            numbering,
            outline_level,
            ..ParagraphProperties::default()
        };
    }
    if !text.is_empty() {
        let run = Run::new(text.to_string());
        let run_id = run.id;
        store.insert(IRNode::Run(run));
        paragraph.runs.push(run_id);
    }
    let para_id = paragraph.id;
    store.insert(IRNode::Paragraph(paragraph));
    para_id
}
