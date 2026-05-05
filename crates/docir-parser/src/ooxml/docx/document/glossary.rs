use super::{parse_block_until, DocxParser};
use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{attr_value, local_name, read_event};
use docir_core::ir::GlossaryEntry;
use docir_core::types::SourceSpan;
use quick_xml::events::Event;
use quick_xml::Reader;

pub(super) fn parse_doc_part(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<GlossaryEntry, ParseError> {
    let mut entry = GlossaryEntry::new();
    entry.span = Some(SourceSpan::new("word/glossary/document.xml"));
    let mut buf = Vec::new();

    loop {
        let event = read_event(reader, &mut buf, "word/glossary/document.xml")?;
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
            Event::End(e) => {
                if local_name(e.name().as_ref()) == b"docPart" {
                    break;
                }
            }
            Event::Eof => break,
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
        let event = read_event(reader, &mut buf, "word/glossary/document.xml")?;
        match event {
            Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"name" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        name = Some(val);
                    }
                }
                b"gallery" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        gallery = Some(val);
                    }
                }
                _ => {}
            },
            Event::End(e) => {
                if local_name(e.name().as_ref()) == b"docPartPr" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok((name, gallery))
}
