use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::read_event;
use docir_core::ir::Section;
use docir_core::types::NodeId;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;

use super::paragraph::{parse_paragraph, parse_paragraph_simple};
use super::table::parse_table;
use super::{apply_section_refs, parse_revision_block, parse_sdt, DocxParser, SdtMode};
use docir_core::ir::RevisionType;

pub(crate) fn parse_body_sections(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<Vec<Section>, ParseError> {
    let mut sections: Vec<Section> = Vec::new();
    let mut current = Section::new();
    let mut buf = Vec::new();

    loop {
        match read_event(reader, &mut buf, "word/document.xml")? {
            Event::Start(e) => match e.name().as_ref() {
                b"w:p" => {
                    let para = parse_paragraph(parser, reader, rels, header_footer_map)?;
                    current.content.push(para.id);
                    if let Some(section_ref) = para.section_ref {
                        current.headers = section_ref.headers;
                        current.footers = section_ref.footers;
                        current.properties = section_ref.properties;
                        sections.push(current);
                        current = Section::new();
                    }
                }
                b"w:tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    current.content.push(table_id);
                }
                b"w:sdt" => {
                    let sdt_id = parse_sdt(parser, reader, rels, SdtMode::Block)?;
                    current.content.push(sdt_id);
                }
                b"w:sectPr" => {
                    let section_ref = apply_section_refs(reader, header_footer_map)?;
                    current.headers = section_ref.headers;
                    current.footers = section_ref.footers;
                    current.properties = section_ref.properties;
                    sections.push(current);
                    current = Section::new();
                }
                b"w:ins" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::Insert)?;
                    current.content.push(rev_id);
                }
                b"w:del" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::Delete)?;
                    current.content.push(rev_id);
                }
                b"w:moveFrom" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::MoveFrom)?;
                    current.content.push(rev_id);
                }
                b"w:moveTo" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::MoveTo)?;
                    current.content.push(rev_id);
                }
                b"w:pPrChange" | b"w:rPrChange" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::FormatChange)?;
                    current.content.push(rev_id);
                }
                _ => {}
            },
            Event::End(e) => {
                if e.name().as_ref() == b"w:body" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if !current.content.is_empty()
        || !current.headers.is_empty()
        || !current.footers.is_empty()
        || current.properties.page_width.is_some()
        || current.properties.page_height.is_some()
        || current.properties.orientation.is_some()
        || current.properties.margins.is_some()
        || current.properties.columns.is_some()
        || sections.is_empty()
    {
        sections.push(current);
    }

    Ok(sections)
}

pub(crate) fn parse_block_until(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    end_tag: &[u8],
) -> Result<Vec<NodeId>, ParseError> {
    let mut content = Vec::new();
    let mut buf = Vec::new();

    loop {
        match read_event(reader, &mut buf, "word/document.xml")? {
            Event::Start(e) => match e.name().as_ref() {
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
            },
            Event::End(e) => {
                if e.name().as_ref() == end_tag {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(content)
}
