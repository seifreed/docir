use super::super::super::{drawing::parse_drawing, support::parse_vml_pict, table::parse_table};
use crate::error::ParseError;
use crate::ooxml::docx::document::span_from_reader;
use crate::ooxml::docx::document::{
    insert_note_reference, parse_paragraph_simple, CommentRangeEnd, CommentRangeStart,
    CommentReference, Run, RunProperties,
};
use crate::ooxml::docx::DocxParser;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{attr_value, XmlScanControl};
use docir_core::ir::{Revision, RevisionType};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

pub(crate) struct RunParse {
    pub(crate) run_id: NodeId,
    pub(crate) text: String,
    pub(crate) has_instr: bool,
    pub(crate) field_char: Option<String>,
    pub(crate) embedded: Vec<NodeId>,
}

#[derive(Debug, Clone, Copy)]
pub(crate) enum SdtMode {
    Block,
    Inline,
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
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"w:r"),
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
    match start.name().as_ref() {
        b"w:rPr" => {
            super::parse_run_properties(reader, &mut state.props)?;
        }
        b"w:drawing" => {
            if let Some(shape_id) = parse_drawing(parser, reader, rels)? {
                state.embedded.push(shape_id);
            }
        }
        b"w:pict" => {
            if let Some(shape_id) = parse_vml_pict(parser, reader, rels)? {
                state.embedded.push(shape_id);
            }
        }
        b"w:footnoteReference" => push_note_reference_if_present(
            parser,
            reader,
            start,
            docir_core::ir::FieldKind::FootnoteRef,
            &mut state.embedded,
        ),
        b"w:endnoteReference" => push_note_reference_if_present(
            parser,
            reader,
            start,
            docir_core::ir::FieldKind::EndnoteRef,
            &mut state.embedded,
        ),
        b"w:fldChar" => {
            state.field_char = attr_value(start, b"w:fldCharType");
        }
        b"w:t" | b"w:instrText" | b"w:delText" => {
            let content = reader.read_text(start.name()).unwrap_or_default();
            if start.name().as_ref() == b"w:instrText" {
                state.has_instr = true;
            }
            state.text.push_str(&content);
        }
        b"w:tab" => state.text.push('\t'),
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
    if start.name().as_ref() == b"w:tab" {
        state.text.push('\t');
        return Ok(());
    }
    if start.name().as_ref() == b"w:fldChar" {
        state.field_char = attr_value(start, b"w:fldCharType");
        return Ok(());
    }
    if start.name().as_ref() == b"w:footnoteReference" {
        push_note_reference_if_present(
            parser,
            reader,
            start,
            docir_core::ir::FieldKind::FootnoteRef,
            &mut state.embedded,
        );
        return Ok(());
    }
    if start.name().as_ref() == b"w:endnoteReference" {
        push_note_reference_if_present(
            parser,
            reader,
            start,
            docir_core::ir::FieldKind::EndnoteRef,
            &mut state.embedded,
        );
        return Ok(());
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
                        e.name().as_ref(),
                        b"w:ins" | b"w:del" | b"w:moveFrom" | b"w:moveTo" | b"w:pPrChange"
                            | b"w:rPrChange"
                    )
            )
        },
        |reader, event| {
            if let Event::Start(e) = event {
                match mode {
                    RevisionParseMode::Inline => {
                        if e.name().as_ref() == b"w:r" {
                            let run = parse_run(parser, reader, rels)?;
                            revision.content.push(run.run_id);
                            revision.content.extend(run.embedded);
                        }
                    }
                    RevisionParseMode::Block => match e.name().as_ref() {
                        b"w:p" => {
                            let para_id = parse_paragraph_simple(parser, reader, rels)?;
                            revision.content.push(para_id);
                        }
                        b"w:tbl" => {
                            let table_id = parse_table(parser, reader, rels)?;
                            revision.content.push(table_id);
                        }
                        b"w:r" => {
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

pub(crate) fn parse_sdt(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    mode: SdtMode,
) -> Result<NodeId, ParseError> {
    let mut control = docir_core::ir::ContentControl::new();
    control.span = Some(span_from_reader(reader, super::DOC_XML_PATH));

    let mut buf = Vec::new();
    super::scan_docx_xml_events_until_end_with_handlers(
        reader,
        &mut buf,
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"w:sdt"),
        |reader, start| {
            match start.name().as_ref() {
                b"w:sdtPr" => {
                    parse_sdt_properties(reader, &mut control)?;
                }
                b"w:sdtContent" => {
                    let content = match mode {
                        SdtMode::Block => parse_sdt_content_block(parser, reader, rels)?,
                        SdtMode::Inline => parse_sdt_content_inline(parser, reader, rels)?,
                    };
                    control.content.extend(content);
                }
                _ => {}
            }
            Ok(())
        },
        |_reader, _event| Ok(()),
        |_reader, _event| Ok(()),
    )?;

    let id = control.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::ContentControl(control));
    Ok(id)
}

fn parse_sdt_properties(
    reader: &mut Reader<&[u8]>,
    control: &mut docir_core::ir::ContentControl,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    #[cfg(test)]
    eprintln!("parse_sdt_properties start");
    super::scan_docx_xml_events_until_end(
        reader,
        &mut buf,
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"w:sdtPr"),
        |_reader, event| {
            match event {
                Event::Start(e) | Event::Empty(e) => match e.name().as_ref() {
                    b"w:tag" => {
                        if let Some(val) = attr_value(e, b"w:val") {
                            control.tag = Some(val);
                            #[cfg(test)]
                            eprintln!("parse_sdt_properties tag={:?}", control.tag);
                        }
                    }
                    b"w:alias" => {
                        if let Some(val) = attr_value(e, b"w:val") {
                            control.alias = Some(val);
                            #[cfg(test)]
                            eprintln!("parse_sdt_properties alias={:?}", control.alias);
                        }
                    }
                    b"w:id" => {
                        if let Some(val) = attr_value(e, b"w:val") {
                            control.sdt_id = Some(val);
                        }
                    }
                    b"w:comboBox" => control.control_type = Some("comboBox".to_string()),
                    b"w:dropDownList" => control.control_type = Some("dropDownList".to_string()),
                    b"w:date" => control.control_type = Some("date".to_string()),
                    b"w:checkbox" => control.control_type = Some("checkbox".to_string()),
                    b"w:text" => control.control_type = Some("text".to_string()),
                    b"w:picture" => control.control_type = Some("picture".to_string()),
                    b"w:dataBinding" => {
                        control.data_binding_xpath = attr_value(e, b"w:xpath");
                        control.data_binding_store_item_id = attr_value(e, b"w:storeItemID");
                        control.data_binding_prefix_mappings = attr_value(e, b"w:prefixMappings");
                    }
                    _ => {}
                },
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        },
    )?;
    #[cfg(test)]
    eprintln!(
        "parse_sdt_properties done alias={:?} tag={:?}",
        control.alias, control.tag
    );
    Ok(())
}

fn parse_sdt_content_block(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<Vec<NodeId>, ParseError> {
    let mut content = Vec::new();
    let mut buf = Vec::new();
    super::scan_docx_xml_events_until_end_start_only(
        reader,
        &mut buf,
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"w:sdtContent"),
        |reader, start| {
            match start.name().as_ref() {
                b"w:p" => {
                    let para_id = parse_paragraph_simple(parser, reader, rels)?;
                    content.push(para_id);
                }
                b"w:tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    content.push(table_id);
                }
                b"w:sdt" => {
                    let sdt_id = parse_sdt(parser, reader, rels, SdtMode::Block)?;
                    content.push(sdt_id);
                }
                _ => {}
            }
            Ok(())
        },
        |_reader, _event| Ok(()),
    )?;
    Ok(content)
}

fn parse_sdt_content_inline(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<Vec<NodeId>, ParseError> {
    let mut runs = Vec::new();
    let mut buf = Vec::new();
    super::scan_docx_xml_events_until_end_start_only(
        reader,
        &mut buf,
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"w:sdtContent"),
        |reader, start| {
            handle_sdt_content_inline_start(parser, reader, rels, start, &mut runs)?;
            Ok(())
        },
        |_reader, _event| Ok(()),
    )?;
    Ok(runs)
}

fn handle_sdt_content_inline_start(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart<'_>,
    runs: &mut Vec<NodeId>,
) -> Result<(), ParseError> {
    match start.name().as_ref() {
        b"w:r" => {
            let run = parse_run(parser, reader, rels)?;
            runs.push(run.run_id);
            runs.extend(run.embedded);
        }
        b"w:hyperlink" => {
            let link_id = super::parse_hyperlink(parser, reader, rels, start)?;
            runs.push(link_id);
        }
        b"w:fldSimple" => {
            let instr = attr_value(start, b"w:instr");
            let field_id = super::parse_field(parser, reader, instr)?;
            runs.push(field_id);
        }
        b"w:commentRangeStart" => {
            if let Some(node_id) = insert_comment_range_start(parser, start) {
                runs.push(node_id);
            }
        }
        b"w:commentRangeEnd" => {
            if let Some(node_id) = insert_comment_range_end(parser, start) {
                runs.push(node_id);
            }
        }
        b"w:commentReference" => {
            if let Some(node_id) = insert_comment_reference(parser, start) {
                runs.push(node_id);
            }
        }
        b"w:bookmarkStart" => {
            if let Some(node_id) = insert_bookmark_start(parser, start) {
                runs.push(node_id);
            }
        }
        b"w:bookmarkEnd" => {
            if let Some(node_id) = insert_bookmark_end(parser, start) {
                runs.push(node_id);
            }
        }
        b"w:ins" => {
            let rev_id = parse_revision_inline(parser, reader, rels, start, RevisionType::Insert)?;
            runs.push(rev_id);
        }
        b"w:del" => {
            let rev_id = parse_revision_inline(parser, reader, rels, start, RevisionType::Delete)?;
            runs.push(rev_id);
        }
        _ => {}
    }
    Ok(())
}

fn insert_comment_range_start(parser: &mut DocxParser, start: &BytesStart<'_>) -> Option<NodeId> {
    let cid = attr_value(start, b"w:id")?;
    let mut node = CommentRangeStart::new(cid);
    node.span = Some(SourceSpan::new(super::DOC_XML_PATH));
    let node_id = node.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::CommentRangeStart(node));
    Some(node_id)
}

fn insert_comment_range_end(parser: &mut DocxParser, start: &BytesStart<'_>) -> Option<NodeId> {
    let cid = attr_value(start, b"w:id")?;
    let mut node = CommentRangeEnd::new(cid);
    node.span = Some(SourceSpan::new(super::DOC_XML_PATH));
    let node_id = node.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::CommentRangeEnd(node));
    Some(node_id)
}

fn insert_comment_reference(parser: &mut DocxParser, start: &BytesStart<'_>) -> Option<NodeId> {
    let cid = attr_value(start, b"w:id")?;
    let mut node = CommentReference::new(cid);
    node.span = Some(SourceSpan::new(super::DOC_XML_PATH));
    let node_id = node.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::CommentReference(node));
    Some(node_id)
}

fn insert_bookmark_start(parser: &mut DocxParser, start: &BytesStart<'_>) -> Option<NodeId> {
    let bm_id = attr_value(start, b"w:id")?;
    let mut bm = docir_core::ir::BookmarkStart::new(bm_id);
    bm.name = attr_value(start, b"w:name");
    let node_id = bm.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::BookmarkStart(bm));
    Some(node_id)
}

fn insert_bookmark_end(parser: &mut DocxParser, start: &BytesStart<'_>) -> Option<NodeId> {
    let bm_id = attr_value(start, b"w:id")?;
    let bm = docir_core::ir::BookmarkEnd::new(bm_id);
    let node_id = bm.id;
    parser.store.insert(docir_core::ir::IRNode::BookmarkEnd(bm));
    Some(node_id)
}
