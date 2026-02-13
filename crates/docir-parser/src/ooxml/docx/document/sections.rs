use crate::error::ParseError;
use crate::xml_utils::{attr_value, xml_error};
use docir_core::ir::{
    LineNumberRestart, PageMargins, PageOrientation, SectionProperties, SectionType,
};
use docir_core::types::NodeId;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::HashMap;

#[derive(Debug, Clone)]
pub(super) struct SectionRef {
    pub(super) headers: Vec<NodeId>,
    pub(super) footers: Vec<NodeId>,
    pub(super) properties: SectionProperties,
}

pub(super) fn apply_section_refs(
    reader: &mut Reader<&[u8]>,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<SectionRef, ParseError> {
    let mut headers = Vec::new();
    let mut footers = Vec::new();
    let mut properties = SectionProperties::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:headerReference" | b"w:footerReference" => {
                    apply_section_header_footer(&e, header_footer_map, &mut headers, &mut footers);
                }
                b"w:pgSz" => apply_section_page_size(&e, &mut properties),
                b"w:pgMar" => apply_section_margins(&e, &mut properties),
                b"w:cols" => apply_section_columns(&e, &mut properties),
                b"w:type" => apply_section_type(&e, &mut properties),
                b"w:titlePg" => apply_section_title_page(&e, &mut properties),
                b"w:pgNumType" => apply_section_page_numbering(&e, &mut properties),
                b"w:lnNumType" | b"w:lineNumberType" => {
                    apply_section_line_numbering(&e, &mut properties)
                }
                b"w:pgBorders" => apply_section_page_borders(reader, &mut properties)?,
                b"w:textDirection" => apply_section_text_direction(&e, &mut properties),
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:sectPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(SectionRef {
        headers,
        footers,
        properties,
    })
}

fn apply_section_header_footer(
    e: &BytesStart<'_>,
    header_footer_map: Option<&HashMap<String, NodeId>>,
    headers: &mut Vec<NodeId>,
    footers: &mut Vec<NodeId>,
) {
    let Some(map) = header_footer_map else {
        return;
    };
    let Some(id) = attr_value(e, b"r:id") else {
        return;
    };
    let Some(node_id) = map.get(&id) else {
        return;
    };
    if e.name().as_ref() == b"w:headerReference" {
        headers.push(*node_id);
    } else {
        footers.push(*node_id);
    }
}

fn apply_section_page_size(e: &BytesStart<'_>, properties: &mut SectionProperties) {
    if let Some(val) = attr_value(e, b"w:w").and_then(|v| v.parse().ok()) {
        properties.page_width = Some(val);
    }
    if let Some(val) = attr_value(e, b"w:h").and_then(|v| v.parse().ok()) {
        properties.page_height = Some(val);
    }
    if let Some(val) = attr_value(e, b"w:orient") {
        properties.orientation = match val.as_str() {
            "landscape" => Some(PageOrientation::Landscape),
            "portrait" => Some(PageOrientation::Portrait),
            _ => None,
        };
    }
}

fn apply_section_margins(e: &BytesStart<'_>, properties: &mut SectionProperties) {
    let mut margins = properties.margins.take().unwrap_or(PageMargins {
        top: 0,
        bottom: 0,
        left: 0,
        right: 0,
        header: None,
        footer: None,
        gutter: None,
    });
    if let Some(val) = attr_value(e, b"w:top").and_then(|v| v.parse().ok()) {
        margins.top = val;
    }
    if let Some(val) = attr_value(e, b"w:bottom").and_then(|v| v.parse().ok()) {
        margins.bottom = val;
    }
    if let Some(val) = attr_value(e, b"w:left").and_then(|v| v.parse().ok()) {
        margins.left = val;
    }
    if let Some(val) = attr_value(e, b"w:right").and_then(|v| v.parse().ok()) {
        margins.right = val;
    }
    if let Some(val) = attr_value(e, b"w:header").and_then(|v| v.parse().ok()) {
        margins.header = Some(val);
    }
    if let Some(val) = attr_value(e, b"w:footer").and_then(|v| v.parse().ok()) {
        margins.footer = Some(val);
    }
    if let Some(val) = attr_value(e, b"w:gutter").and_then(|v| v.parse().ok()) {
        margins.gutter = Some(val);
    }
    properties.margins = Some(margins);
}

fn apply_section_columns(e: &BytesStart<'_>, properties: &mut SectionProperties) {
    if let Some(val) = attr_value(e, b"w:num").and_then(|v| v.parse().ok()) {
        properties.columns = Some(val);
    }
    if let Some(val) = attr_value(e, b"w:space").and_then(|v| v.parse().ok()) {
        properties.column_spacing = Some(val);
    }
    if let Some(val) = attr_value(e, b"w:sep") {
        properties.column_separator = Some(val == "1" || val.eq_ignore_ascii_case("true"));
    }
}

fn apply_section_type(e: &BytesStart<'_>, properties: &mut SectionProperties) {
    if let Some(val) = attr_value(e, b"w:val") {
        properties.section_type = match val.as_str() {
            "continuous" => Some(SectionType::Continuous),
            "evenPage" => Some(SectionType::EvenPage),
            "oddPage" => Some(SectionType::OddPage),
            "nextPage" => Some(SectionType::NextPage),
            _ => None,
        };
    }
}

fn apply_section_title_page(e: &BytesStart<'_>, properties: &mut SectionProperties) {
    properties.title_page = Some(super::bool_from_val(e));
}

fn apply_section_page_numbering(e: &BytesStart<'_>, properties: &mut SectionProperties) {
    let mut numbering = properties.page_numbering.take().unwrap_or_default();
    if let Some(val) = attr_value(e, b"w:start").and_then(|v| v.parse().ok()) {
        numbering.start = Some(val);
    }
    if let Some(val) = attr_value(e, b"w:fmt") {
        numbering.format = Some(val);
    }
    properties.page_numbering = Some(numbering);
}

fn apply_section_line_numbering(e: &BytesStart<'_>, properties: &mut SectionProperties) {
    let mut numbering = properties.line_numbering.take().unwrap_or_default();
    if let Some(val) = attr_value(e, b"w:start").and_then(|v| v.parse().ok()) {
        numbering.start = Some(val);
    }
    if let Some(val) = attr_value(e, b"w:countBy").and_then(|v| v.parse().ok()) {
        numbering.count_by = Some(val);
    }
    if let Some(val) = attr_value(e, b"w:distance").and_then(|v| v.parse().ok()) {
        numbering.distance = Some(val);
    }
    if let Some(val) = attr_value(e, b"w:restart") {
        numbering.restart = match val.as_str() {
            "newPage" => Some(LineNumberRestart::NewPage),
            "newSection" => Some(LineNumberRestart::NewSection),
            "continuous" => Some(LineNumberRestart::Continuous),
            _ => None,
        };
    }
    properties.line_numbering = Some(numbering);
}

fn apply_section_page_borders(
    reader: &mut Reader<&[u8]>,
    properties: &mut SectionProperties,
) -> Result<(), ParseError> {
    if let Some(borders) = super::parse_page_borders(reader)? {
        properties.page_borders = Some(borders);
    }
    Ok(())
}

fn apply_section_text_direction(e: &BytesStart<'_>, properties: &mut SectionProperties) {
    if let Some(val) = attr_value(e, b"w:val") {
        properties.text_direction = Some(val);
    }
}
