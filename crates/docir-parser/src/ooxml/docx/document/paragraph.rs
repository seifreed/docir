use super::inline::{RunParse, SdtMode, parse_revision_inline, parse_run, parse_sdt};
use super::{DocxParser, ParagraphParse, parse_field, parse_hyperlink, span_from_reader};
#[path = "paragraph_props.rs"]
mod paragraph_props;
use super::SectionRef;
use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{local_name, try_attr_value, xml_error};
use docir_core::ir::RevisionType;
use docir_core::ir::{CommentRangeEnd, CommentRangeStart, CommentReference, Field, Paragraph};
use docir_core::types::NodeId;
pub(super) use paragraph_props::parse_paragraph_properties;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;

pub(super) fn parse_paragraph(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<ParagraphParse, ParseError> {
    let mut state = ParagraphParseState::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => handle_paragraph_start_event(
                parser,
                reader,
                rels,
                header_footer_map,
                &mut state,
                &e,
            )?,
            Ok(Event::Empty(e)) => handle_paragraph_empty_event(parser, reader, &mut state, &e)?,
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"p" => {
                break;
            }
            Ok(Event::Eof) => {
                return Err(xml_error(
                    "word/document.xml",
                    "unexpected end of paragraph",
                ));
            }
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    state.para.span = Some(span_from_reader(reader, "word/document.xml"));
    let id = state.para.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::Paragraph(state.para));
    Ok(ParagraphParse {
        id,
        section_ref: state.section_ref,
    })
}

pub(super) fn parse_paragraph_simple(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<NodeId, ParseError> {
    Ok(parse_paragraph(parser, reader, rels, None)?.id)
}

struct FieldState {
    active: bool,
    instr_done: bool,
    instr: String,
    runs: Vec<NodeId>,
}

impl FieldState {
    fn new() -> Self {
        Self {
            active: false,
            instr_done: false,
            instr: String::new(),
            runs: Vec::new(),
        }
    }

    fn start(&mut self) {
        self.active = true;
        self.instr_done = false;
        self.instr.clear();
        self.runs.clear();
    }

    fn separate(&mut self) {
        self.instr_done = true;
    }

    fn finish(&mut self) {
        self.active = false;
        self.instr_done = false;
        self.instr.clear();
        self.runs.clear();
    }
}

struct ParagraphParseState {
    para: Paragraph,
    field_state: FieldState,
    section_ref: Option<SectionRef>,
}

impl ParagraphParseState {
    fn new() -> Self {
        Self {
            para: Paragraph::new(),
            field_state: FieldState::new(),
            section_ref: None,
        }
    }
}

fn update_field_from_run(run: &RunParse, run_id: NodeId, state: &mut FieldState) {
    if state.active {
        state.runs.push(run_id);
        if run.has_instr && !state.instr_done {
            state.instr.push_str(&run.text);
        }
    }
}

fn handle_paragraph_start_event(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    header_footer_map: Option<&HashMap<String, NodeId>>,
    state: &mut ParagraphParseState,
    element: &BytesStart<'_>,
) -> Result<(), ParseError> {
    if local_name(element.name().as_ref()) == b"pPr" {
        state.section_ref = parse_paragraph_properties(reader, &mut state.para, header_footer_map)?;
        return Ok(());
    }

    if handle_inline_start(parser, reader, rels, state, element)? {
        return Ok(());
    }
    if handle_comment_start(parser, reader, state, element)? {
        return Ok(());
    }
    if handle_revision_start(parser, reader, rels, state, element)? {
        return Ok(());
    }
    if handle_bookmark_start(parser, state, element)? {
        return Ok(());
    }

    Ok(())
}

fn handle_inline_start(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    state: &mut ParagraphParseState,
    element: &BytesStart<'_>,
) -> Result<bool, ParseError> {
    match local_name(element.name().as_ref()) {
        b"r" => {
            let run = parse_run(parser, reader, rels)?;
            let run_id = run.run_id;
            state.para.runs.push(run_id);
            for emb in &run.embedded {
                state.para.runs.push(*emb);
            }
            update_field_from_run(&run, run_id, &mut state.field_state);
            handle_field_char(
                parser,
                &mut state.para,
                &mut state.field_state,
                run.field_char.as_deref(),
            );
            Ok(true)
        }
        b"hyperlink" => {
            let link_id = parse_hyperlink(parser, reader, rels, element)?;
            state.para.runs.push(link_id);
            Ok(true)
        }
        b"sdt" => {
            let sdt_id = parse_sdt(parser, reader, rels, SdtMode::Inline)?;
            state.para.runs.push(sdt_id);
            Ok(true)
        }
        b"fldSimple" => {
            let instr = try_attr_value(element, b"w:instr", "word/document.xml")?;
            let field_id = parse_field(parser, reader, instr)?;
            state.para.runs.push(field_id);
            Ok(true)
        }
        _ => Ok(false),
    }
}

fn handle_comment_start(
    parser: &mut DocxParser,
    reader: &Reader<&[u8]>,
    state: &mut ParagraphParseState,
    element: &BytesStart<'_>,
) -> Result<bool, ParseError> {
    let kind = match local_name(element.name().as_ref()) {
        b"commentRangeStart" => CommentNodeKind::RangeStart,
        b"commentRangeEnd" => CommentNodeKind::RangeEnd,
        b"commentReference" => CommentNodeKind::Reference,
        _ => return Ok(false),
    };
    insert_comment_node(parser, reader, &mut state.para, element, kind)?;
    Ok(true)
}

fn handle_revision_start(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    state: &mut ParagraphParseState,
    element: &BytesStart<'_>,
) -> Result<bool, ParseError> {
    let revision_type = match local_name(element.name().as_ref()) {
        b"ins" => RevisionType::Insert,
        b"del" => RevisionType::Delete,
        b"moveFrom" => RevisionType::MoveFrom,
        b"moveTo" => RevisionType::MoveTo,
        b"pPrChange" | b"rPrChange" => RevisionType::FormatChange,
        _ => return Ok(false),
    };
    push_revision_inline(
        parser,
        reader,
        rels,
        &mut state.para,
        element,
        revision_type,
    )?;
    Ok(true)
}

fn handle_bookmark_start(
    parser: &mut DocxParser,
    state: &mut ParagraphParseState,
    element: &BytesStart<'_>,
) -> Result<bool, ParseError> {
    match local_name(element.name().as_ref()) {
        b"bookmarkStart" => {
            insert_bookmark_start(parser, &mut state.para, element)?;
            Ok(true)
        }
        b"bookmarkEnd" => {
            insert_bookmark_end(parser, &mut state.para, element)?;
            Ok(true)
        }
        _ => Ok(false),
    }
}

fn handle_paragraph_empty_event(
    parser: &mut DocxParser,
    reader: &Reader<&[u8]>,
    state: &mut ParagraphParseState,
    element: &BytesStart<'_>,
) -> Result<(), ParseError> {
    match local_name(element.name().as_ref()) {
        b"commentRangeStart" => insert_comment_node(
            parser,
            reader,
            &mut state.para,
            element,
            CommentNodeKind::RangeStart,
        )?,
        b"commentRangeEnd" => insert_comment_node(
            parser,
            reader,
            &mut state.para,
            element,
            CommentNodeKind::RangeEnd,
        )?,
        b"commentReference" => insert_comment_node(
            parser,
            reader,
            &mut state.para,
            element,
            CommentNodeKind::Reference,
        )?,
        b"bookmarkStart" => insert_bookmark_start(parser, &mut state.para, element)?,
        b"bookmarkEnd" => insert_bookmark_end(parser, &mut state.para, element)?,
        _ => {}
    }
    Ok(())
}

fn push_revision_inline(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    para: &mut Paragraph,
    element: &BytesStart<'_>,
    revision_type: RevisionType,
) -> Result<(), ParseError> {
    let rev_id = parse_revision_inline(parser, reader, rels, element, revision_type)?;
    para.runs.push(rev_id);
    Ok(())
}

enum CommentNodeKind {
    RangeStart,
    RangeEnd,
    Reference,
}

fn insert_comment_node(
    parser: &mut DocxParser,
    reader: &Reader<&[u8]>,
    para: &mut Paragraph,
    element: &BytesStart<'_>,
    kind: CommentNodeKind,
) -> Result<(), ParseError> {
    if let Some(cid) = try_attr_value(element, b"w:id", "word/document.xml")? {
        let span = span_from_reader(reader, "word/document.xml");
        let (node, node_id) = match kind {
            CommentNodeKind::RangeStart => {
                let mut node = CommentRangeStart::new(cid);
                node.span = Some(span);
                let node_id = node.id;
                (docir_core::ir::IRNode::CommentRangeStart(node), node_id)
            }
            CommentNodeKind::RangeEnd => {
                let mut node = CommentRangeEnd::new(cid);
                node.span = Some(span);
                let node_id = node.id;
                (docir_core::ir::IRNode::CommentRangeEnd(node), node_id)
            }
            CommentNodeKind::Reference => {
                let mut node = CommentReference::new(cid);
                node.span = Some(span);
                let node_id = node.id;
                (docir_core::ir::IRNode::CommentReference(node), node_id)
            }
        };
        parser.store.insert(node);
        para.runs.push(node_id);
    }
    Ok(())
}

fn insert_bookmark_start(
    parser: &mut DocxParser,
    para: &mut Paragraph,
    element: &BytesStart<'_>,
) -> Result<(), ParseError> {
    if let Some(bm_id) = try_attr_value(element, b"w:id", "word/document.xml")? {
        let mut bm = docir_core::ir::BookmarkStart::new(bm_id);
        bm.name = try_attr_value(element, b"w:name", "word/document.xml")?;
        bm.col_first = bookmark_column_attr(element, b"w:colFirst")?;
        bm.col_last = bookmark_column_attr(element, b"w:colLast")?;
        let bm_id = bm.id;
        parser
            .store
            .insert(docir_core::ir::IRNode::BookmarkStart(bm));
        para.runs.push(bm_id);
    }
    Ok(())
}

fn bookmark_column_attr(element: &BytesStart<'_>, name: &[u8]) -> Result<Option<u32>, ParseError> {
    try_attr_value(element, name, "word/document.xml")?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|err| xml_error("word/document.xml", err))
        })
        .transpose()
}

fn insert_bookmark_end(
    parser: &mut DocxParser,
    para: &mut Paragraph,
    element: &BytesStart<'_>,
) -> Result<(), ParseError> {
    if let Some(bm_id) = try_attr_value(element, b"w:id", "word/document.xml")? {
        let bm = docir_core::ir::BookmarkEnd::new(bm_id);
        let bm_id = bm.id;
        parser.store.insert(docir_core::ir::IRNode::BookmarkEnd(bm));
        para.runs.push(bm_id);
    }
    Ok(())
}

fn handle_field_char(
    parser: &mut DocxParser,
    para: &mut Paragraph,
    state: &mut FieldState,
    char_type: Option<&str>,
) {
    match char_type {
        Some("begin") => state.start(),
        Some("separate") => state.separate(),
        Some("end") => {
            if state.active {
                let instr = if state.instr.trim().is_empty() {
                    None
                } else {
                    Some(state.instr.trim().to_string())
                };
                let mut field = Field::new(instr);
                field.runs = state.runs.clone();
                let field_id = field.id;
                parser.store.insert(docir_core::ir::IRNode::Field(field));
                para.runs.push(field_id);
            }
            state.finish();
        }
        _ => {}
    }
}

#[cfg(test)]
#[path = "paragraph_tests.rs"]
mod tests;
