use crate::error::ParseError;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::borrow::Cow;
use std::fmt::Display;
use std::io::BufRead;

/// Converts an XML attribute's value bytes to a lossy UTF-8 string.
/// Replaces the verbose `String::from_utf8_lossy(&attr.value)` pattern in attribute iteration loops.
pub(crate) fn lossy_attr_value<'a>(
    attr: &'a quick_xml::events::attributes::Attribute<'a>,
) -> Cow<'a, str> {
    String::from_utf8_lossy(&attr.value)
}

pub(crate) fn attr_value(e: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == name {
            return Some(lossy_attr_value(&attr).to_string());
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

pub(crate) fn attr_bool_like(raw: &[u8]) -> bool {
    raw == b"1" || raw.eq_ignore_ascii_case(b"true")
}

pub(crate) fn attr_u32_from_bytes(e: &BytesStart<'_>, name: &[u8]) -> Option<u32> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == name {
            let value = std::str::from_utf8(attr.value.as_ref()).ok()?;
            return value.parse::<u32>().ok();
        }
    }
    None
}

pub(crate) fn attr_u64_from_bytes(e: &BytesStart<'_>, name: &[u8]) -> Option<u64> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == name {
            let value = std::str::from_utf8(attr.value.as_ref()).ok()?;
            return value.parse::<u64>().ok();
        }
    }
    None
}

pub(crate) fn attr_each<'a, F>(element: &'a BytesStart<'a>, mut on_attr: F)
where
    F: for<'b> FnMut(&'b [u8], &'b [u8]),
{
    for attr in element.attributes().flatten() {
        let key = local_name(attr.key.as_ref());
        on_attr(key, attr.value.as_ref());
    }
}

pub(crate) fn attr_f64(e: &BytesStart<'_>, name: &[u8]) -> Option<f64> {
    attr_value(e, name).and_then(|v| v.parse::<f64>().ok())
}

pub(crate) fn is_end_event_local(event: &Event<'_>, name: &[u8]) -> bool {
    matches!(event, Event::End(e) if local_name(e.name().as_ref()) == name)
}

pub(crate) fn dispatch_start_or_empty<'e, R, F>(
    reader: &mut Reader<R>,
    event: &'e Event<'e>,
    mut on_event: F,
) -> Result<bool, ParseError>
where
    R: BufRead,
    F: FnMut(&mut Reader<R>, &'e BytesStart<'e>, bool) -> Result<(), ParseError>,
{
    match event {
        Event::Start(start) => {
            on_event(reader, start, true)?;
            Ok(true)
        }
        Event::Empty(empty) => {
            on_event(reader, empty, false)?;
            Ok(true)
        }
        _ => Ok(false),
    }
}

pub(crate) fn attr_value_by_suffix(e: &BytesStart<'_>, suffixes: &[&[u8]]) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        for suffix in suffixes {
            if key.ends_with(suffix) {
                if let Ok(value) = attr.unescape_value() {
                    return Some(value.to_string());
                }
                return Some(lossy_attr_value(&attr).to_string());
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
    cfg.check_end_names = true;
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum XmlScanControl {
    Continue,
    Break,
}

const MAX_XML_NESTING_DEPTH: usize = 512;

pub(crate) fn scan_xml_events<R, F>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
    file: &str,
    mut on_event: F,
) -> Result<(), ParseError>
where
    R: BufRead,
    F: for<'a> FnMut(Event<'a>) -> Result<XmlScanControl, ParseError>,
{
    reader.config_mut().check_end_names = true;
    let mut start_elements: Vec<Vec<u8>> = Vec::new();

    loop {
        let event = match reader.read_event_into(buf) {
            Ok(event) => event,
            Err(err) => return Err(xml_error(file, err)),
        };
        let is_eof = matches!(&event, Event::Eof);
        let mut push_open: Option<Vec<u8>> = None;
        let mut check_close: Option<Vec<u8>> = None;

        match &event {
            Event::Start(e) => push_open = Some(e.name().as_ref().to_vec()),
            Event::End(e) => check_close = Some(e.name().as_ref().to_vec()),
            Event::Empty(_) | Event::Text(_) | Event::Comment(_) => {}
            _ => {}
        }

        if on_event(event)? == XmlScanControl::Break {
            break;
        }

        if let Some(name) = push_open {
            if start_elements.len() >= MAX_XML_NESTING_DEPTH {
                return Err(xml_error(
                    file,
                    format!("XML nesting depth exceeded maximum ({MAX_XML_NESTING_DEPTH})"),
                ));
            }
            start_elements.push(name);
        }
        if let Some(name) = check_close {
            match start_elements.pop() {
                Some(start) if start != name => {
                    return Err(xml_error(
                        file,
                        format!(
                            "unexpected end tag: {}",
                            String::from_utf8_lossy(name.as_slice())
                        ),
                    ));
                }
                Some(_) => {}
                None => {
                    return Err(xml_error(
                        file,
                        format!(
                            "unexpected end tag: {}",
                            String::from_utf8_lossy(name.as_slice())
                        ),
                    ));
                }
            }
        }

        if is_eof && !start_elements.is_empty() {
            return Err(xml_error(
                file,
                format!(
                    "unexpected end tag: {}",
                    String::from_utf8_lossy(start_elements[start_elements.len() - 1].as_slice()),
                ),
            ));
        }
        if is_eof {
            break;
        }
        buf.clear();
    }
    Ok(())
}

pub(crate) fn scan_xml_events_with_reader<R, F>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
    file: &str,
    mut on_event: F,
) -> Result<(), ParseError>
where
    R: BufRead,
    F: for<'a> FnMut(&mut Reader<R>, Event<'a>) -> Result<XmlScanControl, ParseError>,
{
    reader.config_mut().check_end_names = true;

    loop {
        match reader.read_event_into(buf) {
            Ok(event) => {
                let is_eof = matches!(&event, Event::Eof);
                if on_event(reader, event)? == XmlScanControl::Break {
                    break;
                }
                if is_eof {
                    break;
                }
            }
            Err(err) => return Err(xml_error(file, err)),
        }
        buf.clear();
    }
    Ok(())
}

pub(crate) fn scan_xml_events_until_end<R, B, F>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
    file: &str,
    mut is_end: B,
    mut on_event: F,
) -> Result<(), ParseError>
where
    R: BufRead,
    B: for<'e> FnMut(&Event<'e>) -> bool,
    F: for<'e> FnMut(&mut Reader<R>, &'e Event<'e>) -> Result<XmlScanControl, ParseError>,
{
    scan_xml_events_with_reader(reader, buf, file, |reader, event| {
        if is_end(&event) {
            return Ok(XmlScanControl::Break);
        }
        on_event(reader, &event)
    })
}

pub(crate) fn scan_xml_events_until_end_dispatch<R, B, FStart, FEmpty, FOther>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
    file: &str,
    mut is_end: B,
    mut on_start: FStart,
    mut on_empty: FEmpty,
    mut on_other: FOther,
) -> Result<(), ParseError>
where
    R: BufRead,
    B: for<'e> FnMut(&Event<'e>) -> bool,
    FStart: for<'e> FnMut(&mut Reader<R>, &BytesStart<'e>) -> Result<(), ParseError>,
    FEmpty: for<'e> FnMut(&mut Reader<R>, &BytesStart<'e>) -> Result<(), ParseError>,
    FOther: for<'e> FnMut(&mut Reader<R>, &Event<'e>) -> Result<(), ParseError>,
{
    scan_xml_events_with_reader(reader, buf, file, |reader, event| {
        if is_end(&event) {
            return Ok(XmlScanControl::Break);
        }

        match event {
            Event::Start(start) => on_start(reader, &start)?,
            Event::Empty(empty) => on_empty(reader, &empty)?,
            _ => on_other(reader, &event)?,
        }

        Ok(XmlScanControl::Continue)
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scan_xml_events_stops_on_signal() {
        let mut reader = Reader::from_str("<root><a></a><b/></root>");
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut seen = Vec::new();

        scan_xml_events(&mut reader, &mut buf, "scan.xml", |event| {
            match event {
                Event::Start(start) => {
                    seen.push(String::from_utf8_lossy(start.name().as_ref()).to_string());
                }
                Event::End(end) if end.name().as_ref() == b"a" => {
                    return Ok(XmlScanControl::Break);
                }
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        })
        .expect("scan should not fail");

        assert_eq!(seen, vec!["root".to_string(), "a".to_string()]);
    }

    #[test]
    fn scan_xml_events_maps_parse_error() {
        let mut reader = Reader::from_str("<root><a>");
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let err = scan_xml_events(&mut reader, &mut buf, "broken.xml", |_| {
            Ok(XmlScanControl::Continue)
        })
        .expect_err("malformed xml should fail");
        assert!(format!("{err}").contains("broken.xml"));
    }

    #[test]
    fn scan_xml_events_until_end_breaks_on_predicate() {
        let mut reader = Reader::from_str("<root><a><b/></a><c/></root>");
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut names = Vec::new();

        scan_xml_events_until_end(
            &mut reader,
            &mut buf,
            "scan_until.xml",
            |event| matches!(event, Event::End(e) if e.name().as_ref() == b"a"),
            |_, event| {
                if let Event::Start(start) = event {
                    names.push(String::from_utf8_lossy(start.name().as_ref()).to_string());
                }
                Ok(XmlScanControl::Continue)
            },
        )
        .expect("scan should stop on predicate");

        assert_eq!(names, vec!["root".to_string(), "a".to_string()]);
    }
}
