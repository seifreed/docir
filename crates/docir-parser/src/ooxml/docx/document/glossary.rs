use super::{DocxParser, parse_block_until};
use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{local_name, read_event, try_attr_value, xml_error};
use docir_core::ir::GlossaryEntry;
use docir_core::types::SourceSpan;
use quick_xml::Reader;
use quick_xml::events::Event;

const GLOSSARY_PATH: &str = "word/glossary/document.xml";

pub(super) fn parse_doc_part(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<GlossaryEntry, ParseError> {
    let mut entry = GlossaryEntry::new();
    entry.span = Some(SourceSpan::new(GLOSSARY_PATH));
    let mut buf = Vec::new();

    loop {
        let event = read_event(reader, &mut buf, GLOSSARY_PATH)?;
        match event {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"docPartPr" => {
                    let (name, gallery) = parse_doc_part_pr(reader)?;
                    entry.name = name;
                    entry.gallery = gallery;
                }
                b"docPartBody" => {
                    let content = parse_block_until(parser, reader, rels, b"docPartBody")?;
                    entry.content.extend(content);
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"docPart" => {
                break;
            }
            Event::Eof => {
                return Err(xml_error(GLOSSARY_PATH, "unexpected end of docPart"));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(entry)
}

fn parse_doc_part_pr(
    reader: &mut Reader<&[u8]>,
) -> Result<(Option<String>, Option<String>), ParseError> {
    let mut name = None;
    let mut gallery = None;
    let mut buf = Vec::new();

    loop {
        let event = read_event(reader, &mut buf, GLOSSARY_PATH)?;
        match event {
            Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"name" => {
                    if let Some(val) = try_attr_value(&e, b"w:val", GLOSSARY_PATH)? {
                        name = Some(val);
                    }
                }
                b"gallery" => {
                    if let Some(val) = try_attr_value(&e, b"w:val", GLOSSARY_PATH)? {
                        gallery = Some(val);
                    }
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"docPartPr" => {
                break;
            }
            Event::Eof => {
                return Err(xml_error(GLOSSARY_PATH, "unexpected end of docPartPr"));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok((name, gallery))
}
