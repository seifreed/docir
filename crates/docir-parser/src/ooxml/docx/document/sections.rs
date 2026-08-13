use crate::error::ParseError;
use crate::xml_utils::{local_name, try_attr_value, try_attr_value_by_suffix, xml_error};
use docir_core::ir::{
    LineNumberRestart, PageMargins, PageOrientation, SectionProperties, SectionType,
};
use docir_core::types::NodeId;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;

const DOC_PATH: &str = "word/document.xml";

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
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                b"headerReference" | b"footerReference" => {
                    apply_section_header_footer(&e, header_footer_map, &mut headers, &mut footers)?;
                }
                b"pgSz" => apply_section_page_size(&e, &mut properties)?,
                b"pgMar" => apply_section_margins(&e, &mut properties)?,
                b"cols" => apply_section_columns(&e, &mut properties)?,
                b"type" => apply_section_type(&e, &mut properties)?,
                b"titlePg" => apply_section_title_page(&e, &mut properties)?,
                b"pgNumType" => apply_section_page_numbering(&e, &mut properties)?,
                b"lnNumType" | b"lineNumberType" => {
                    apply_section_line_numbering(&e, &mut properties)?
                }
                b"pgBorders" => apply_section_page_borders(reader, &mut properties)?,
                b"textDirection" => apply_section_text_direction(&e, &mut properties)?,
                _ => {}
            },
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"sectPr" => {
                break;
            }
            Ok(Event::Eof) => {
                return Err(xml_error(DOC_PATH, "unexpected end of sectPr"));
            }
            Err(e) => {
                return Err(xml_error(DOC_PATH, e));
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
) -> Result<(), ParseError> {
    let Some(map) = header_footer_map else {
        return Ok(());
    };
    let Some(id) = try_attr_value_by_suffix(e, &[b":id"], DOC_PATH)? else {
        return Ok(());
    };
    let Some(node_id) = map.get(&id) else {
        return Ok(());
    };
    if local_name(e.name().as_ref()) == b"headerReference" {
        headers.push(*node_id);
    } else {
        footers.push(*node_id);
    }
    Ok(())
}

fn apply_section_page_size(
    e: &BytesStart<'_>,
    properties: &mut SectionProperties,
) -> Result<(), ParseError> {
    if let Some(val) = u32_attr(e, b"w:w")? {
        properties.page_width = Some(val);
    }
    if let Some(val) = u32_attr(e, b"w:h")? {
        properties.page_height = Some(val);
    }
    if let Some(val) = try_attr_value(e, b"w:orient", DOC_PATH)? {
        properties.orientation = match val.as_str() {
            "landscape" => Some(PageOrientation::Landscape),
            "portrait" => Some(PageOrientation::Portrait),
            _ => None,
        };
    }
    Ok(())
}

fn apply_section_margins(
    e: &BytesStart<'_>,
    properties: &mut SectionProperties,
) -> Result<(), ParseError> {
    let mut margins = properties.margins.take().unwrap_or(PageMargins {
        top: 0,
        bottom: 0,
        left: 0,
        right: 0,
        header: None,
        footer: None,
        gutter: None,
    });
    if let Some(val) = u32_attr(e, b"w:top")? {
        margins.top = val;
    }
    if let Some(val) = u32_attr(e, b"w:bottom")? {
        margins.bottom = val;
    }
    if let Some(val) = u32_attr(e, b"w:left")? {
        margins.left = val;
    }
    if let Some(val) = u32_attr(e, b"w:right")? {
        margins.right = val;
    }
    if let Some(val) = u32_attr(e, b"w:header")? {
        margins.header = Some(val);
    }
    if let Some(val) = u32_attr(e, b"w:footer")? {
        margins.footer = Some(val);
    }
    if let Some(val) = u32_attr(e, b"w:gutter")? {
        margins.gutter = Some(val);
    }
    properties.margins = Some(margins);
    Ok(())
}

fn apply_section_columns(
    e: &BytesStart<'_>,
    properties: &mut SectionProperties,
) -> Result<(), ParseError> {
    if let Some(val) = u32_attr(e, b"w:num")? {
        properties.columns = Some(val);
    }
    if let Some(val) = u32_attr(e, b"w:space")? {
        properties.column_spacing = Some(val);
    }
    if let Some(val) = try_attr_value(e, b"w:sep", DOC_PATH)? {
        properties.column_separator = Some(val == "1" || val.eq_ignore_ascii_case("true"));
    }
    Ok(())
}

fn apply_section_type(
    e: &BytesStart<'_>,
    properties: &mut SectionProperties,
) -> Result<(), ParseError> {
    if let Some(val) = try_attr_value(e, b"w:val", DOC_PATH)? {
        properties.section_type = match val.as_str() {
            "continuous" => Some(SectionType::Continuous),
            "evenPage" => Some(SectionType::EvenPage),
            "oddPage" => Some(SectionType::OddPage),
            "nextPage" => Some(SectionType::NextPage),
            _ => None,
        };
    }
    Ok(())
}

fn apply_section_title_page(
    e: &BytesStart<'_>,
    properties: &mut SectionProperties,
) -> Result<(), ParseError> {
    try_attr_value(e, b"w:val", DOC_PATH)?;
    properties.title_page = Some(super::bool_from_val(e, DOC_PATH)?);
    Ok(())
}

fn apply_section_page_numbering(
    e: &BytesStart<'_>,
    properties: &mut SectionProperties,
) -> Result<(), ParseError> {
    let mut numbering = properties.page_numbering.take().unwrap_or_default();
    if let Some(val) = u32_attr(e, b"w:start")? {
        numbering.start = Some(val);
    }
    if let Some(val) = try_attr_value(e, b"w:fmt", DOC_PATH)? {
        numbering.format = Some(val);
    }
    properties.page_numbering = Some(numbering);
    Ok(())
}

fn apply_section_line_numbering(
    e: &BytesStart<'_>,
    properties: &mut SectionProperties,
) -> Result<(), ParseError> {
    let mut numbering = properties.line_numbering.take().unwrap_or_default();
    if let Some(val) = u32_attr(e, b"w:start")? {
        numbering.start = Some(val);
    }
    if let Some(val) = u32_attr(e, b"w:countBy")? {
        numbering.count_by = Some(val);
    }
    if let Some(val) = u32_attr(e, b"w:distance")? {
        numbering.distance = Some(val);
    }
    if let Some(val) = try_attr_value(e, b"w:restart", DOC_PATH)? {
        numbering.restart = match val.as_str() {
            "newPage" => Some(LineNumberRestart::NewPage),
            "newSection" => Some(LineNumberRestart::NewSection),
            "continuous" => Some(LineNumberRestart::Continuous),
            _ => None,
        };
    }
    properties.line_numbering = Some(numbering);
    Ok(())
}

fn u32_attr(e: &BytesStart<'_>, name: &[u8]) -> Result<Option<u32>, ParseError> {
    try_attr_value(e, name, DOC_PATH)?
        .map(|value| value.parse::<u32>().map_err(|err| xml_error(DOC_PATH, err)))
        .transpose()
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

fn apply_section_text_direction(
    e: &BytesStart<'_>,
    properties: &mut SectionProperties,
) -> Result<(), ParseError> {
    if let Some(val) = try_attr_value(e, b"w:val", DOC_PATH)? {
        properties.text_direction = Some(val);
    }
    Ok(())
}
