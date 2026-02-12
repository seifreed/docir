use super::DocxParser;
use crate::error::ParseError;
use crate::xml_utils::attr_value;
use docir_core::ir::{Paragraph, RunProperties, Style, StyleSet, StyleType};
use docir_core::types::NodeId;
use quick_xml::events::Event;
use quick_xml::Reader;

use super::paragraph::parse_paragraph_properties;
use super::table::parse_table_properties;

impl DocxParser {
    pub fn parse_styles(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let mut styles = StyleSet::new();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut current: Option<Style> = None;
        let mut in_name = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"w:style" => {
                        let style_id = attr_value(&e, b"w:styleId").unwrap_or_default();
                        let mut style = Style {
                            style_id,
                            name: None,
                            style_type: StyleType::Other,
                            based_on: None,
                            next: None,
                            is_default: attr_value(&e, b"w:default")
                                .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
                                .unwrap_or(false),
                            run_props: None,
                            paragraph_props: None,
                            table_props: None,
                        };
                        if let Some(t) = attr_value(&e, b"w:type") {
                            style.style_type = match t.as_str() {
                                "paragraph" => StyleType::Paragraph,
                                "character" => StyleType::Character,
                                "table" => StyleType::Table,
                                "numbering" => StyleType::Numbering,
                                _ => StyleType::Other,
                            };
                        }
                        current = Some(style);
                    }
                    b"w:name" => {
                        in_name = true;
                    }
                    b"w:rPr" => {
                        let mut props = RunProperties::default();
                        super::parse_run_properties(&mut reader, &mut props)?;
                        if let Some(style) = current.as_mut() {
                            style.run_props = Some(super::style_run_from_run_props(props));
                        }
                    }
                    b"w:pPr" => {
                        let mut para = Paragraph::new();
                        let _ = parse_paragraph_properties(&mut reader, &mut para, None)?;
                        if let Some(style) = current.as_mut() {
                            style.paragraph_props =
                                Some(super::style_paragraph_from_paragraph_props(para.properties));
                        }
                    }
                    b"w:tblPr" => {
                        if let Some(style) = current.as_mut() {
                            let mut props = docir_core::ir::TableProperties::default();
                            parse_table_properties(&mut reader, &mut props)?;
                            style.table_props = Some(props);
                        }
                    }
                    b"w:basedOn" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(style) = current.as_mut() {
                                style.based_on = Some(val);
                            }
                        }
                    }
                    b"w:next" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(style) = current.as_mut() {
                                style.next = Some(val);
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => match e.name().as_ref() {
                    b"w:name" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(style) = current.as_mut() {
                                style.name = Some(val);
                            }
                        }
                    }
                    b"w:basedOn" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(style) = current.as_mut() {
                                style.based_on = Some(val);
                            }
                        }
                    }
                    b"w:next" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(style) = current.as_mut() {
                                style.next = Some(val);
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::Text(e)) => {
                    if in_name {
                        if let Some(style) = current.as_mut() {
                            style.name = Some(e.unescape().unwrap_or_default().to_string());
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"w:name" {
                        in_name = false;
                    } else if e.name().as_ref() == b"w:style" {
                        if let Some(style) = current.take() {
                            styles.styles.push(style);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: "word/styles.xml".to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        let id = styles.id;
        self.store.insert(docir_core::ir::IRNode::StyleSet(styles));
        Ok(id)
    }

    pub fn parse_styles_with_effects(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let id = self.parse_styles(xml)?;
        if let Some(docir_core::ir::IRNode::StyleSet(set)) = self.store.get_mut(id) {
            set.with_effects = true;
        }
        Ok(id)
    }
}
