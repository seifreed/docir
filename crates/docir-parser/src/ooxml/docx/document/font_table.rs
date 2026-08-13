use crate::error::ParseError;
use crate::ooxml::docx::document::DocxParser;
use crate::xml_utils::{local_name, reader_from_str, try_attr_value, xml_error};
use docir_core::ir::{FontEntry, FontTable};
use docir_core::types::NodeId;
use quick_xml::events::Event;

const FONT_TABLE_PATH: &str = "word/fontTable.xml";

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
                Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"font" => {
                    let name =
                        try_attr_value(&e, b"w:name", FONT_TABLE_PATH)?.ok_or_else(|| {
                            ParseError::InvalidStructure(
                                "word/fontTable.xml font is missing w:name".to_string(),
                            )
                        })?;
                    current = Some(FontEntry {
                        name,
                        alt_name: None,
                        charset: None,
                        family: None,
                        panose: None,
                    });
                }
                Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                    b"altName" => {
                        if let Some(val) = try_attr_value(&e, b"w:val", FONT_TABLE_PATH)?
                            && let Some(font) = current.as_mut()
                        {
                            font.alt_name = Some(val);
                        }
                    }
                    b"charset" => {
                        if let Some(val) = charset_attr(&e)?
                            && let Some(font) = current.as_mut()
                        {
                            font.charset = Some(val);
                        }
                    }
                    b"family" => {
                        if let Some(val) = try_attr_value(&e, b"w:val", FONT_TABLE_PATH)?
                            && let Some(font) = current.as_mut()
                        {
                            font.family = Some(val);
                        }
                    }
                    b"panose1" => {
                        if let Some(val) = try_attr_value(&e, b"w:val", FONT_TABLE_PATH)?
                            && let Some(font) = current.as_mut()
                        {
                            font.panose = Some(val);
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => {
                    if local_name(e.name().as_ref()) == b"font"
                        && let Some(font) = current.take()
                    {
                        table.fonts.push(font);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(xml_error(FONT_TABLE_PATH, e));
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

fn charset_attr(e: &quick_xml::events::BytesStart<'_>) -> Result<Option<u32>, ParseError> {
    try_attr_value(e, b"w:val", FONT_TABLE_PATH)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|err| xml_error(FONT_TABLE_PATH, err))
        })
        .transpose()
}
