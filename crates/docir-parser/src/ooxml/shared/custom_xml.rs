use crate::error::ParseError;
use crate::xml_utils::xml_error;
use docir_core::ir::CustomXmlPart;
use docir_core::types::SourceSpan;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashSet;

pub fn parse_custom_xml_part(
    xml: &str,
    path: &str,
    size_bytes: u64,
) -> Result<CustomXmlPart, ParseError> {
    let mut part = CustomXmlPart::new(path, size_bytes);
    part.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut namespaces: HashSet<String> = HashSet::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                part.root_element = Some(String::from_utf8_lossy(e.name().as_ref()).to_string());
                for attr in e.attributes().flatten() {
                    let key = String::from_utf8_lossy(attr.key.as_ref());
                    if key.starts_with("xmlns") {
                        namespaces.insert(String::from_utf8_lossy(&attr.value).to_string());
                    }
                }
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    part.namespaces = namespaces.into_iter().collect();
    Ok(part)
}
