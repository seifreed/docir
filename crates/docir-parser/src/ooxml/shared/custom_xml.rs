use crate::error::ParseError;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::xml_error;
use crate::xml_utils::{track_xml_document_event, visit_attributes};
use docir_core::ir::CustomXmlPart;
use docir_core::types::SourceSpan;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashSet;

/// Public API entrypoint: parse_custom_xml_part.
pub fn parse_custom_xml_part(
    xml: &str,
    path: &str,
    size_bytes: u64,
) -> Result<CustomXmlPart, ParseError> {
    let mut part = CustomXmlPart::new(path, size_bytes);
    part.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;

    let mut buf = Vec::new();
    let mut namespaces: HashSet<String> = HashSet::new();
    let mut depth = 0usize;
    let mut root_closed = false;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error(path, err))?;
        if track_xml_document_event(&event, &mut depth, &mut root_closed, path)? {
            break;
        }
        match event {
            Event::Start(e) | Event::Empty(e) if part.root_element.is_none() => {
                part.root_element = Some(String::from_utf8_lossy(e.name().as_ref()).to_string());
                visit_attributes(&e, path, |attr| {
                    let key = String::from_utf8_lossy(attr.key.as_ref());
                    if key.starts_with("xmlns") {
                        namespaces.insert(lossy_attr_value(attr).to_string());
                    }
                })?;
            }
            _ => {}
        }
        buf.clear();
    }

    part.namespaces = namespaces.into_iter().collect();
    Ok(part)
}
