use crate::error::ParseError;
use crate::ooxml::docx::document::DocxParser;
use crate::xml_utils::{attr_value, reader_from_str, xml_error};
use docir_core::ir::{FontEntry, FontTable};
use docir_core::types::NodeId;
use quick_xml::events::Event;

impl DocxParser {
    /// Public API entrypoint: parse_font_table.
    pub fn parse_font_table(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        self.parse_font_table_impl(xml)
    }

    fn parse_font_table_impl(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let mut table = FontTable::new();
        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();
        let mut current: Option<FontEntry> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"w:font" {
                        let name = attr_value(&e, b"w:name").unwrap_or_default();
                        current = Some(FontEntry {
                            name,
                            alt_name: None,
                            charset: None,
                            family: None,
                            panose: None,
                        });
                    }
                }
                Ok(Event::Empty(e)) => match e.name().as_ref() {
                    b"w:altName" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(font) = current.as_mut() {
                                font.alt_name = Some(val);
                            }
                        }
                    }
                    b"w:charset" => {
                        if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                            if let Some(font) = current.as_mut() {
                                font.charset = Some(val);
                            }
                        }
                    }
                    b"w:family" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(font) = current.as_mut() {
                                font.family = Some(val);
                            }
                        }
                    }
                    b"w:panose1" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(font) = current.as_mut() {
                                font.panose = Some(val);
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"w:font" {
                        if let Some(font) = current.take() {
                            table.fonts.push(font);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(xml_error("word/fontTable.xml", e));
                }
                _ => {}
            }
            buf.clear();
        }

        let id = table.id;
        self.store.insert(docir_core::ir::IRNode::FontTable(table));
        Ok(id)
    }
}
