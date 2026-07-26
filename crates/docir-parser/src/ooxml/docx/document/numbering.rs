use super::DocxParser;
use crate::error::ParseError;
use crate::xml_utils::{local_name, try_attr_value, xml_error};
use docir_core::ir::{NumberingLevel, NumberingSet, Paragraph, RunProperties, TextAlignment};
use docir_core::types::NodeId;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use super::paragraph::parse_paragraph_properties;

const NUMBERING_PATH: &str = "word/numbering.xml";

impl DocxParser {
    /// Public API entrypoint: parse_numbering.
    pub fn parse_numbering(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let mut set = NumberingSet::new();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut current_abs: Option<u32> = None;
        let mut current_levels: Vec<NumberingLevel> = Vec::new();
        let mut current_level: Option<NumberingLevel> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => handle_numbering_start(
                    &mut reader,
                    &e,
                    &mut set,
                    &mut current_abs,
                    &mut current_levels,
                    &mut current_level,
                )?,
                Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                    b"lvl" => {
                        handle_numbering_start(
                            &mut reader,
                            &e,
                            &mut set,
                            &mut current_abs,
                            &mut current_levels,
                            &mut current_level,
                        )?;
                        if let Some(level) = current_level.take() {
                            current_levels.push(level);
                        }
                    }
                    _ => handle_level_value_attrs(&e, current_level.as_mut())?,
                },
                Ok(Event::End(e)) => match local_name(e.name().as_ref()) {
                    b"lvl" => {
                        if let Some(level) = current_level.take() {
                            current_levels.push(level);
                        }
                    }
                    b"abstractNum" => {
                        if let Some(abs_id) = current_abs.take() {
                            set.abstract_nums.push(docir_core::ir::AbstractNum {
                                abstract_id: abs_id,
                                levels: current_levels.clone(),
                            });
                        }
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(xml_error(NUMBERING_PATH, e));
                }
                _ => {}
            }
            buf.clear();
        }

        let id = set.id;
        self.store.insert(docir_core::ir::IRNode::NumberingSet(set));
        Ok(id)
    }
}

fn handle_numbering_start(
    reader: &mut Reader<&[u8]>,
    event: &BytesStart<'_>,
    set: &mut NumberingSet,
    current_abs: &mut Option<u32>,
    current_levels: &mut Vec<NumberingLevel>,
    current_level: &mut Option<NumberingLevel>,
) -> Result<(), ParseError> {
    match local_name(event.name().as_ref()) {
        b"abstractNum" => {
            *current_abs = u32_attr(event, b"w:abstractNumId")?;
            current_levels.clear();
        }
        b"lvl" => {
            let lvl = u32_attr(event, b"w:ilvl")?.unwrap_or(0);
            *current_level = Some(NumberingLevel {
                level: lvl,
                format: None,
                text: None,
                start: None,
                alignment: None,
                suffix: None,
                paragraph_props: None,
                run_props: None,
            });
        }
        b"numFmt" | b"lvlText" | b"start" | b"lvlJc" | b"suff" => {
            handle_level_value_attrs(event, current_level.as_mut())?;
        }
        b"pPr" => {
            let mut para = Paragraph::new();
            let _ = parse_paragraph_properties(reader, &mut para, None)?;
            if let Some(level) = current_level.as_mut() {
                level.paragraph_props =
                    Some(super::style_paragraph_from_paragraph_props(para.properties));
            }
        }
        b"rPr" => {
            let mut props = RunProperties::default();
            super::parse_run_properties(reader, &mut props)?;
            if let Some(level) = current_level.as_mut() {
                level.run_props = Some(super::style_run_from_run_props(props));
            }
        }
        b"num" => {
            let num_id = u32_attr(event, b"w:numId")?.unwrap_or(0);
            let abstract_id = super::parse_num_abstract_id(reader)?;
            set.nums.push(docir_core::ir::NumInstance {
                num_id,
                abstract_id,
            });
        }
        _ => {}
    }
    Ok(())
}

fn handle_level_value_attrs(
    event: &BytesStart<'_>,
    level: Option<&mut NumberingLevel>,
) -> Result<(), ParseError> {
    let Some(level) = level else {
        return Ok(());
    };
    match local_name(event.name().as_ref()) {
        b"numFmt" => {
            if let Some(val) = try_attr_value(event, b"w:val", NUMBERING_PATH)? {
                level.format = Some(val);
            }
        }
        b"lvlText" => {
            if let Some(val) = try_attr_value(event, b"w:val", NUMBERING_PATH)? {
                level.text = Some(val);
            }
        }
        b"start" => {
            if let Some(val) = u32_attr(event, b"w:val")? {
                level.start = Some(val);
            }
        }
        b"lvlJc" => {
            if let Some(val) = try_attr_value(event, b"w:val", NUMBERING_PATH)? {
                level.alignment = match val.as_str() {
                    "center" => Some(TextAlignment::Center),
                    "right" => Some(TextAlignment::Right),
                    "justify" => Some(TextAlignment::Justify),
                    "distribute" => Some(TextAlignment::Distribute),
                    _ => Some(TextAlignment::Left),
                };
            }
        }
        b"suff" => {
            if let Some(val) = try_attr_value(event, b"w:val", NUMBERING_PATH)? {
                level.suffix = Some(val);
            }
        }
        _ => {}
    }
    Ok(())
}

fn u32_attr(event: &BytesStart<'_>, name: &[u8]) -> Result<Option<u32>, ParseError> {
    try_attr_value(event, name, NUMBERING_PATH)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|err| xml_error(NUMBERING_PATH, err))
        })
        .transpose()
}
