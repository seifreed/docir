use super::{parse_block_until, DocxParser};
use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{attr_value, read_event};
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
            Event::Start(e) => match e.name().as_ref() {
                b"w:docPartPr" => {
                    let (name, gallery) = parse_doc_part_pr(reader)?;
                    entry.name = name;
                    entry.gallery = gallery;
                }
                b"w:docPartBody" => {
                    let content = parse_block_until(parser, reader, rels, b"w:docPartBody")?;
                    entry.content.extend(content);
                }
                _ => {}
            },
            Event::End(e) => {
                if e.name().as_ref() == b"w:docPart" {
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
            Event::Start(e) | Event::Empty(e) => match e.name().as_ref() {
                b"w:name" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        name = Some(val);
                    }
                }
                b"w:gallery" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        gallery = Some(val);
                    }
                }
                _ => {}
            },
            Event::End(e) => {
                if e.name().as_ref() == b"w:docPartPr" {
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
