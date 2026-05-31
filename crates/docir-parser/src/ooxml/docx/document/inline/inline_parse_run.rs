use super::super::super::{drawing::parse_drawing, support::parse_vml_pict, table::parse_table};
use crate::error::ParseError;
use crate::ooxml::docx::DocxParser;
use crate::ooxml::docx::document::span_from_reader;
use crate::ooxml::docx::document::{
    Run, RunProperties, insert_note_reference, parse_paragraph_simple,
};
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{XmlScanControl, attr_value, local_name, xml_error};
use docir_core::ir::{Revision, RevisionType};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

pub(crate) struct RunParse {
    pub(crate) run_id: NodeId,
    pub(crate) text: String,
    pub(crate) has_instr: bool,
    pub(crate) field_char: Option<String>,
    pub(crate) embedded: Vec<NodeId>,
}

pub(crate) fn parse_run(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<RunParse, ParseError> {
    let mut state = RunParseState::default();
    let mut buf = Vec::new();

    super::scan_docx_xml_events_until_end(
        reader,
        &mut buf,
        |event| matches!(event, Event::End(e) if local_name(e.name().as_ref()) == b"r"),
        |reader, event| {
            match event {
                Event::Start(start) => {
                    handle_run_start_event(parser, reader, rels, start, &mut state)?;
                }
                Event::Empty(start) => {
                    handle_run_empty_event(parser, reader, start, &mut state)?;
                }
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        },
    )?;

    let mut run = Run::new(state.text.clone());
    run.properties = state.props;
    run.span = Some(span_from_reader(reader, super::DOC_XML_PATH));
    let run_text = run.text.clone();
    let id = run.id;
    parser.store.insert(docir_core::ir::IRNode::Run(run));
    Ok(RunParse {
        run_id: id,
        text: run_text,
        has_instr: state.has_instr,
        field_char: state.field_char,
        embedded: state.embedded,
    })
}

#[derive(Debug, Default)]
struct RunParseState {
    text: String,
    props: RunProperties,
    has_instr: bool,
    field_char: Option<String>,
    embedded: Vec<NodeId>,
}

fn handle_run_start_event(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart<'_>,
    state: &mut RunParseState,
) -> Result<(), ParseError> {
    match local_name(start.name().as_ref()) {
        b"rPr" => {
            super::parse_run_properties(reader, &mut state.props)?;
        }
        b"drawing" => {
            if let Some(shape_id) = parse_drawing(parser, reader, rels)? {
                state.embedded.push(shape_id);
            }
        }
        b"pict" => {
            if let Some(shape_id) = parse_vml_pict(parser, reader, rels)? {
                state.embedded.push(shape_id);
            }
        }
        b"footnoteReference" => push_note_reference_if_present(
            parser,
            reader,
            start,
            docir_core::ir::FieldKind::FootnoteRef,
            &mut state.embedded,
        ),
        b"endnoteReference" => push_note_reference_if_present(
            parser,
            reader,
            start,
            docir_core::ir::FieldKind::EndnoteRef,
            &mut state.embedded,
        ),
        b"fldChar" => {
            state.field_char = attr_value(start, b"w:fldCharType");
        }
        b"t" | b"instrText" | b"delText" => {
            let content = reader
                .read_text(start.name())
                .map_err(|err| xml_error(super::DOC_XML_PATH, err))?;
            if local_name(start.name().as_ref()) == b"instrText" {
                state.has_instr = true;
            }
            state.text.push_str(&content);
        }
        b"tab" => state.text.push('\t'),
        _ => {}
    }
    Ok(())
}

fn handle_run_empty_event(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    start: &BytesStart<'_>,
    state: &mut RunParseState,
) -> Result<(), ParseError> {
    match local_name(start.name().as_ref()) {
        b"tab" => {
            state.text.push('\t');
        }
        b"fldChar" => {
            state.field_char = attr_value(start, b"w:fldCharType");
        }
        b"footnoteReference" => push_note_reference_if_present(
            parser,
            reader,
            start,
            docir_core::ir::FieldKind::FootnoteRef,
            &mut state.embedded,
        ),
        b"endnoteReference" => push_note_reference_if_present(
            parser,
            reader,
            start,
            docir_core::ir::FieldKind::EndnoteRef,
            &mut state.embedded,
        ),
        _ => {}
    }
    Ok(())
}

fn push_note_reference_if_present(
    parser: &mut DocxParser,
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
    kind: docir_core::ir::FieldKind,
    embedded: &mut Vec<NodeId>,
) {
    if let Some(field_id) = parse_note_reference(parser, reader, start, kind) {
        embedded.push(field_id);
    }
}

fn parse_note_reference(
    parser: &mut DocxParser,
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
    kind: docir_core::ir::FieldKind,
) -> Option<NodeId> {
    let id = attr_value(start, b"w:id")?;
    Some(insert_note_reference(parser, reader, kind, id))
}

pub(crate) fn parse_revision_inline(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart,
    change_type: RevisionType,
) -> Result<NodeId, ParseError> {
    parse_revision(
        parser,
        reader,
        rels,
        start,
        change_type,
        RevisionParseMode::Inline,
    )
}

pub(crate) fn parse_revision_block(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart,
    change_type: RevisionType,
) -> Result<NodeId, ParseError> {
    parse_revision(
        parser,
        reader,
        rels,
        start,
        change_type,
        RevisionParseMode::Block,
    )
}

enum RevisionParseMode {
    Inline,
    Block,
}

fn parse_revision(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart,
    change_type: RevisionType,
    mode: RevisionParseMode,
) -> Result<NodeId, ParseError> {
    let mut revision = Revision::new(change_type);
    revision.revision_id = attr_value(start, b"w:id");
    revision.author = attr_value(start, b"w:author");
    revision.date = attr_value(start, b"w:date");
    revision.span = Some(SourceSpan::new(super::DOC_XML_PATH));

    let mut buf = Vec::new();
    super::scan_docx_xml_events_until_end(
        reader,
        &mut buf,
        |event| {
            matches!(
                event,
                Event::End(e)
                    if matches!(
                        local_name(e.name().as_ref()),
                        b"ins" | b"del" | b"moveFrom" | b"moveTo" | b"pPrChange"
                            | b"rPrChange"
                    )
            )
        },
        |reader, event| {
            if let Event::Start(e) = event {
                match mode {
                    RevisionParseMode::Inline => {
                        if local_name(e.name().as_ref()) == b"r" {
                            let run = parse_run(parser, reader, rels)?;
                            revision.content.push(run.run_id);
                            revision.content.extend(run.embedded);
                        }
                    }
                    RevisionParseMode::Block => match local_name(e.name().as_ref()) {
                        b"p" => {
                            let para_id = parse_paragraph_simple(parser, reader, rels)?;
                            revision.content.push(para_id);
                        }
                        b"tbl" => {
                            let table_id = parse_table(parser, reader, rels)?;
                            revision.content.push(table_id);
                        }
                        b"r" => {
                            let run = parse_run(parser, reader, rels)?;
                            revision.content.push(run.run_id);
                            revision.content.extend(run.embedded);
                        }
                        _ => {}
                    },
                }
            }
            Ok(XmlScanControl::Continue)
        },
    )?;

    let id = revision.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::Revision(revision));
    Ok(id)
}
