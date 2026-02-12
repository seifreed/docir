use crate::error::ParseError;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::fmt::Display;
use std::io::BufRead;

pub(crate) fn attr_value(e: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == name {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

pub(crate) fn attr_u32(e: &BytesStart<'_>, name: &[u8]) -> Option<u32> {
    attr_value(e, name).and_then(|v| v.parse::<u32>().ok())
}

pub(crate) fn attr_bool(e: &BytesStart<'_>, name: &[u8]) -> Option<bool> {
    attr_value(e, name).map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
}

pub(crate) fn attr_f64(e: &BytesStart<'_>, name: &[u8]) -> Option<f64> {
    attr_value(e, name).and_then(|v| v.parse::<f64>().ok())
}

pub(crate) fn attr_value_by_suffix(e: &BytesStart<'_>, suffixes: &[&[u8]]) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        for suffix in suffixes {
            if key.ends_with(suffix) {
                if let Ok(value) = attr.unescape_value() {
                    return Some(value.to_string());
                }
                return Some(String::from_utf8_lossy(&attr.value).to_string());
            }
        }
    }
    None
}

pub(crate) fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(pos) => &name[pos + 1..],
        None => name,
    }
}

pub(crate) fn xml_error(file: &str, err: impl Display) -> ParseError {
    ParseError::Xml {
        file: file.to_string(),
        message: err.to_string(),
    }
}

pub(crate) fn reader_from_str(xml: &str) -> Reader<&[u8]> {
    reader_from_str_with_options(xml, true, false)
}

pub(crate) fn reader_from_str_with_options(
    xml: &str,
    trim_text: bool,
    expand_empty_elements: bool,
) -> Reader<&[u8]> {
    let mut reader = Reader::from_str(xml);
    let cfg = reader.config_mut();
    cfg.trim_text(trim_text);
    cfg.expand_empty_elements = expand_empty_elements;
    reader
}

pub(crate) fn read_event<'a, R: BufRead>(
    reader: &mut Reader<R>,
    buf: &'a mut Vec<u8>,
    file: &str,
) -> Result<Event<'a>, ParseError> {
    reader.read_event_into(buf).map_err(|err| ParseError::Xml {
        file: file.to_string(),
        message: err.to_string(),
    })
}
