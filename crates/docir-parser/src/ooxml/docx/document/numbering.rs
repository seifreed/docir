use super::DocxParser;
use crate::error::ParseError;
use crate::xml_utils::attr_value;
use docir_core::ir::{NumberingLevel, NumberingSet, Paragraph, RunProperties, TextAlignment};
use docir_core::types::NodeId;
use quick_xml::events::Event;
use quick_xml::Reader;

use super::paragraph::parse_paragraph_properties;

impl DocxParser {
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
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"w:abstractNum" => {
                        current_abs =
                            attr_value(&e, b"w:abstractNumId").and_then(|v| v.parse().ok());
                        current_levels.clear();
                    }
                    b"w:lvl" => {
                        let lvl = attr_value(&e, b"w:ilvl")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                        current_level = Some(NumberingLevel {
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
                    b"w:numFmt" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.format = Some(val);
                            }
                        }
                    }
                    b"w:lvlText" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.text = Some(val);
                            }
                        }
                    }
                    b"w:start" => {
                        if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                            if let Some(level) = current_level.as_mut() {
                                level.start = Some(val);
                            }
                        }
                    }
                    b"w:lvlJc" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.alignment = match val.as_str() {
                                    "center" => Some(TextAlignment::Center),
                                    "right" => Some(TextAlignment::Right),
                                    "justify" => Some(TextAlignment::Justify),
                                    "distribute" => Some(TextAlignment::Distribute),
                                    _ => Some(TextAlignment::Left),
                                };
                            }
                        }
                    }
                    b"w:suff" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.suffix = Some(val);
                            }
                        }
                    }
                    b"w:pPr" => {
                        let mut para = Paragraph::new();
                        let _ = parse_paragraph_properties(&mut reader, &mut para, None)?;
                        if let Some(level) = current_level.as_mut() {
                            level.paragraph_props =
                                Some(super::style_paragraph_from_paragraph_props(para.properties));
                        }
                    }
                    b"w:rPr" => {
                        let mut props = RunProperties::default();
                        super::parse_run_properties(&mut reader, &mut props)?;
                        if let Some(level) = current_level.as_mut() {
                            level.run_props = Some(super::style_run_from_run_props(props));
                        }
                    }
                    b"w:num" => {
                        let num_id = attr_value(&e, b"w:numId")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                        let abstract_id = super::parse_num_abstract_id(&mut reader)?;
                        set.nums.push(docir_core::ir::NumInstance {
                            num_id,
                            abstract_id,
                        });
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => match e.name().as_ref() {
                    b"w:numFmt" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.format = Some(val);
                            }
                        }
                    }
                    b"w:lvlText" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.text = Some(val);
                            }
                        }
                    }
                    b"w:start" => {
                        if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                            if let Some(level) = current_level.as_mut() {
                                level.start = Some(val);
                            }
                        }
                    }
                    b"w:lvlJc" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.alignment = match val.as_str() {
                                    "center" => Some(TextAlignment::Center),
                                    "right" => Some(TextAlignment::Right),
                                    "justify" => Some(TextAlignment::Justify),
                                    "distribute" => Some(TextAlignment::Distribute),
                                    _ => Some(TextAlignment::Left),
                                };
                            }
                        }
                    }
                    b"w:suff" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.suffix = Some(val);
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => match e.name().as_ref() {
                    b"w:lvl" => {
                        if let Some(level) = current_level.take() {
                            current_levels.push(level);
                        }
                    }
                    b"w:abstractNum" => {
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
                    return Err(ParseError::Xml {
                        file: "word/numbering.xml".to_string(),
                        message: e.to_string(),
                    });
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
