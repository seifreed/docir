use crate::error::ParseError;
use crate::ooxml::docx::field::parse_field_instruction;
use crate::xml_utils::XmlScanControl;
use crate::xml_utils::{scan_xml_events_until_end, scan_xml_events_until_end_dispatch};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

const DOC_XML_PATH: &str = "word/document.xml";

fn scan_docx_xml_events_until_end<FMatch, FEvent>(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    is_end: FMatch,
    mut on_event: FEvent,
) -> Result<(), ParseError>
where
    FMatch: Fn(&Event<'_>) -> bool,
    FEvent: FnMut(&mut Reader<&[u8]>, &Event<'_>) -> Result<XmlScanControl, ParseError>,
{
    scan_xml_events_until_end(reader, buf, DOC_XML_PATH, is_end, |reader, event| {
        on_event(reader, event)
    })
}

fn scan_docx_xml_events_until_end_with_handlers<FMatch, FStart, FEmpty, FOther>(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    is_end: FMatch,
    on_start: FStart,
    on_empty: FEmpty,
    on_other: FOther,
) -> Result<(), ParseError>
where
    FMatch: Fn(&Event<'_>) -> bool,
    FStart: FnMut(&mut Reader<&[u8]>, &BytesStart<'_>) -> Result<(), ParseError>,
    FEmpty: FnMut(&mut Reader<&[u8]>, &BytesStart<'_>) -> Result<(), ParseError>,
    FOther: FnMut(&mut Reader<&[u8]>, &Event<'_>) -> Result<(), ParseError>,
{
    scan_xml_events_until_end_dispatch(
        reader,
        buf,
        DOC_XML_PATH,
        is_end,
        on_start,
        on_empty,
        on_other,
    )
}

fn scan_docx_xml_events_until_end_start_only<FMatch, FStart, FEmpty>(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    is_end: FMatch,
    mut on_start: FStart,
    mut on_empty: FEmpty,
) -> Result<(), ParseError>
where
    FMatch: Fn(&Event<'_>) -> bool,
    FStart: FnMut(&mut Reader<&[u8]>, &BytesStart<'_>) -> Result<(), ParseError>,
    FEmpty: FnMut(&mut Reader<&[u8]>, &BytesStart<'_>) -> Result<(), ParseError>,
{
    scan_docx_xml_events_until_end_with_handlers(
        reader,
        buf,
        is_end,
        |reader, start| {
            on_start(reader, start)?;
            Ok(())
        },
        |reader, start| {
            on_empty(reader, start)?;
            Ok(())
        },
        |_reader, _event| Ok(()),
    )
}

#[path = "inline_parse_markers.rs"]
mod inline_parse_markers;
#[path = "inline_parse_run.rs"]
mod inline_parse_run;

pub(crate) use inline_parse_markers::*;
pub(crate) use inline_parse_run::*;
