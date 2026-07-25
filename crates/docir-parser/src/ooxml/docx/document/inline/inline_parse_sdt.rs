use super::super::super::{parse_paragraph_simple, table::parse_table};
use super::{parse_field, parse_hyperlink, parse_revision_inline, parse_run};
use crate::error::ParseError;
use crate::ooxml::docx::DocxParser;
use crate::ooxml::docx::document::span_from_reader;
use crate::ooxml::docx::document::{CommentRangeEnd, CommentRangeStart, CommentReference};
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{XmlScanControl, local_name, try_attr_value};
use docir_core::ir::RevisionType;
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

#[derive(Debug, Clone, Copy)]
pub(crate) enum SdtMode {
    Block,
    Inline,
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
        |event| matches!(event, Event::End(e) if local_name(e.name().as_ref()) == b"sdt"),
        |reader, start| {
            match local_name(start.name().as_ref()) {
                b"sdtPr" => {
                    parse_sdt_properties(reader, &mut control)?;
                }
                b"sdtContent" => {
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
    super::scan_docx_xml_events_until_end(
        reader,
        &mut buf,
        |event| matches!(event, Event::End(e) if local_name(e.name().as_ref()) == b"sdtPr"),
        |_reader, event| {
            match event {
                Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()) {
                    b"tag" => {
                        if let Some(val) = try_attr_value(e, b"w:val", super::DOC_XML_PATH)? {
                            control.tag = Some(val);
                        }
                    }
                    b"alias" => {
                        if let Some(val) = try_attr_value(e, b"w:val", super::DOC_XML_PATH)? {
                            control.alias = Some(val);
                        }
                    }
                    b"id" => {
                        if let Some(val) = try_attr_value(e, b"w:val", super::DOC_XML_PATH)? {
                            control.sdt_id = Some(val);
                        }
                    }
                    b"comboBox" => control.control_type = Some("comboBox".to_string()),
                    b"dropDownList" => control.control_type = Some("dropDownList".to_string()),
                    b"date" => control.control_type = Some("date".to_string()),
                    b"checkbox" => control.control_type = Some("checkbox".to_string()),
                    b"text" => control.control_type = Some("text".to_string()),
                    b"picture" => control.control_type = Some("picture".to_string()),
                    b"dataBinding" => {
                        control.data_binding_xpath =
                            try_attr_value(e, b"w:xpath", super::DOC_XML_PATH)?;
                        control.data_binding_store_item_id =
                            try_attr_value(e, b"w:storeItemID", super::DOC_XML_PATH)?;
                        control.data_binding_prefix_mappings =
                            try_attr_value(e, b"w:prefixMappings", super::DOC_XML_PATH)?;
                    }
                    _ => {}
                },
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        },
    )?;
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
        |event| matches!(event, Event::End(e) if local_name(e.name().as_ref()) == b"sdtContent"),
        |reader, start| {
            match local_name(start.name().as_ref()) {
                b"p" => {
                    let para_id = parse_paragraph_simple(parser, reader, rels)?;
                    content.push(para_id);
                }
                b"tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    content.push(table_id);
                }
                b"sdt" => {
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
        |event| matches!(event, Event::End(e) if local_name(e.name().as_ref()) == b"sdtContent"),
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
    match local_name(start.name().as_ref()) {
        b"r" => {
            let run = parse_run(parser, reader, rels)?;
            runs.push(run.run_id);
            runs.extend(run.embedded);
        }
        b"hyperlink" => {
            let link_id = parse_hyperlink(parser, reader, rels, start)?;
            runs.push(link_id);
        }
        b"fldSimple" => {
            let instr = try_attr_value(start, b"w:instr", super::DOC_XML_PATH)?;
            let field_id = parse_field(parser, reader, instr)?;
            runs.push(field_id);
        }
        b"commentRangeStart" => {
            if let Some(node_id) = insert_comment_range_start(parser, start)? {
                runs.push(node_id);
            }
        }
        b"commentRangeEnd" => {
            if let Some(node_id) = insert_comment_range_end(parser, start)? {
                runs.push(node_id);
            }
        }
        b"commentReference" => {
            if let Some(node_id) = insert_comment_reference(parser, start)? {
                runs.push(node_id);
            }
        }
        b"bookmarkStart" => {
            if let Some(node_id) = insert_bookmark_start(parser, start)? {
                runs.push(node_id);
            }
        }
        b"bookmarkEnd" => {
            if let Some(node_id) = insert_bookmark_end(parser, start)? {
                runs.push(node_id);
            }
        }
        b"ins" => {
            let rev_id = parse_revision_inline(parser, reader, rels, start, RevisionType::Insert)?;
            runs.push(rev_id);
        }
        b"del" => {
            let rev_id = parse_revision_inline(parser, reader, rels, start, RevisionType::Delete)?;
            runs.push(rev_id);
        }
        _ => {}
    }
    Ok(())
}

fn insert_comment_range_start(
    parser: &mut DocxParser,
    start: &BytesStart<'_>,
) -> Result<Option<NodeId>, ParseError> {
    let Some(cid) = try_attr_value(start, b"w:id", super::DOC_XML_PATH)? else {
        return Ok(None);
    };
    let mut node = CommentRangeStart::new(cid);
    node.span = Some(SourceSpan::new(super::DOC_XML_PATH));
    let node_id = node.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::CommentRangeStart(node));
    Ok(Some(node_id))
}

fn insert_comment_range_end(
    parser: &mut DocxParser,
    start: &BytesStart<'_>,
) -> Result<Option<NodeId>, ParseError> {
    let Some(cid) = try_attr_value(start, b"w:id", super::DOC_XML_PATH)? else {
        return Ok(None);
    };
    let mut node = CommentRangeEnd::new(cid);
    node.span = Some(SourceSpan::new(super::DOC_XML_PATH));
    let node_id = node.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::CommentRangeEnd(node));
    Ok(Some(node_id))
}

fn insert_comment_reference(
    parser: &mut DocxParser,
    start: &BytesStart<'_>,
) -> Result<Option<NodeId>, ParseError> {
    let Some(cid) = try_attr_value(start, b"w:id", super::DOC_XML_PATH)? else {
        return Ok(None);
    };
    let mut node = CommentReference::new(cid);
    node.span = Some(SourceSpan::new(super::DOC_XML_PATH));
    let node_id = node.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::CommentReference(node));
    Ok(Some(node_id))
}

fn insert_bookmark_start(
    parser: &mut DocxParser,
    start: &BytesStart<'_>,
) -> Result<Option<NodeId>, ParseError> {
    let Some(bm_id) = try_attr_value(start, b"w:id", super::DOC_XML_PATH)? else {
        return Ok(None);
    };
    let mut bm = docir_core::ir::BookmarkStart::new(bm_id);
    bm.name = try_attr_value(start, b"w:name", super::DOC_XML_PATH)?;
    let node_id = bm.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::BookmarkStart(bm));
    Ok(Some(node_id))
}

fn insert_bookmark_end(
    parser: &mut DocxParser,
    start: &BytesStart<'_>,
) -> Result<Option<NodeId>, ParseError> {
    let Some(bm_id) = try_attr_value(start, b"w:id", super::DOC_XML_PATH)? else {
        return Ok(None);
    };
    let bm = docir_core::ir::BookmarkEnd::new(bm_id);
    let node_id = bm.id;
    parser.store.insert(docir_core::ir::IRNode::BookmarkEnd(bm));
    Ok(Some(node_id))
}
