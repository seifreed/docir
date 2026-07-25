use super::super::{SectionRef, apply_section_refs, bool_from_val, parse_border, parse_numbering};
use crate::error::ParseError;
use crate::xml_utils::{local_name, try_attr_value, xml_error};
use docir_core::ir::Paragraph;
use docir_core::ir::{LineSpacingRule, TextAlignment};
use docir_core::types::NodeId;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;

const DOC_PATH: &str = "word/document.xml";

pub(crate) fn alignment_from_val(val: &str) -> TextAlignment {
    match val {
        "center" => TextAlignment::Center,
        "right" => TextAlignment::Right,
        "both" => TextAlignment::Justify,
        "distribute" => TextAlignment::Distribute,
        _ => TextAlignment::Left,
    }
}

pub(crate) fn parse_paragraph_properties(
    reader: &mut Reader<&[u8]>,
    para: &mut Paragraph,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<Option<SectionRef>, ParseError> {
    let mut buf = Vec::new();
    let mut section_ref: Option<SectionRef> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                handle_paragraph_property_event(
                    reader,
                    &e,
                    para,
                    header_footer_map,
                    &mut section_ref,
                )?;
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"pPr" => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(section_ref)
}

fn handle_paragraph_property_event(
    reader: &mut Reader<&[u8]>,
    e: &BytesStart<'_>,
    para: &mut Paragraph,
    header_footer_map: Option<&HashMap<String, NodeId>>,
    section_ref: &mut Option<SectionRef>,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"pStyle" => {
            if let Some(val) = try_attr_value(e, b"w:val", DOC_PATH)? {
                para.style_id = Some(val);
            }
        }
        b"keepNext" => para.properties.keep_next = Some(bool_from_val(e)),
        b"keepLines" => para.properties.keep_lines = Some(bool_from_val(e)),
        b"pageBreakBefore" => para.properties.page_break_before = Some(bool_from_val(e)),
        b"widowControl" => para.properties.widow_control = Some(bool_from_val(e)),
        b"jc" => apply_paragraph_alignment(e, para)?,
        b"ind" => apply_paragraph_indentation(e, para)?,
        b"spacing" => apply_paragraph_spacing(e, para)?,
        b"pBdr" => {
            if let Some(borders) = parse_paragraph_borders(reader)? {
                para.properties.borders = Some(borders);
            }
        }
        b"outlineLvl" => {
            if let Some(val) = try_attr_value(e, b"w:val", DOC_PATH)?.and_then(|v| v.parse().ok()) {
                para.properties.outline_level = Some(val);
            }
        }
        b"numPr" => parse_numbering(reader, &mut para.properties)?,
        b"sectPr" => *section_ref = Some(apply_section_refs(reader, header_footer_map)?),
        _ => {}
    }
    Ok(())
}

fn apply_paragraph_alignment(e: &BytesStart<'_>, para: &mut Paragraph) -> Result<(), ParseError> {
    if let Some(val) = try_attr_value(e, b"w:val", DOC_PATH)? {
        para.properties.alignment = Some(alignment_from_val(val.as_str()));
    }
    Ok(())
}

fn apply_paragraph_indentation(e: &BytesStart<'_>, para: &mut Paragraph) -> Result<(), ParseError> {
    let mut indent = para.properties.indentation.clone().unwrap_or_default();
    if let Some(val) = try_attr_value(e, b"w:left", DOC_PATH)?.and_then(|v| v.parse().ok()) {
        indent.left = Some(val);
    }
    if let Some(val) = try_attr_value(e, b"w:right", DOC_PATH)?.and_then(|v| v.parse().ok()) {
        indent.right = Some(val);
    }
    if let Some(val) = try_attr_value(e, b"w:firstLine", DOC_PATH)?.and_then(|v| v.parse().ok()) {
        indent.first_line = Some(val);
    }
    if let Some(val) = try_attr_value(e, b"w:hanging", DOC_PATH)?.and_then(|v| v.parse().ok()) {
        indent.hanging = Some(val);
    }
    para.properties.indentation = Some(indent);
    Ok(())
}

fn apply_paragraph_spacing(e: &BytesStart<'_>, para: &mut Paragraph) -> Result<(), ParseError> {
    let mut spacing = para.properties.spacing.clone().unwrap_or_default();
    if let Some(val) = try_attr_value(e, b"w:before", DOC_PATH)?.and_then(|v| v.parse().ok()) {
        spacing.before = Some(val);
    }
    if let Some(val) = try_attr_value(e, b"w:after", DOC_PATH)?.and_then(|v| v.parse().ok()) {
        spacing.after = Some(val);
    }
    if let Some(val) = try_attr_value(e, b"w:line", DOC_PATH)?.and_then(|v| v.parse().ok()) {
        spacing.line = Some(val);
    }
    if let Some(val) = try_attr_value(e, b"w:lineRule", DOC_PATH)? {
        spacing.line_rule = match val.as_str() {
            "auto" => Some(LineSpacingRule::Auto),
            "exact" => Some(LineSpacingRule::Exact),
            "atLeast" => Some(LineSpacingRule::AtLeast),
            _ => None,
        };
    }
    para.properties.spacing = Some(spacing);
    Ok(())
}

pub(crate) fn parse_paragraph_borders(
    reader: &mut Reader<&[u8]>,
) -> Result<Option<docir_core::ir::ParagraphBorders>, ParseError> {
    let mut buf = Vec::new();
    let mut borders = docir_core::ir::ParagraphBorders::default();
    let mut has_any = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let border = parse_border(&e);
                if border.is_none() {
                    continue;
                }
                match local_name(e.name().as_ref()) {
                    b"top" => {
                        borders.top = border;
                        has_any = true;
                    }
                    b"bottom" => {
                        borders.bottom = border;
                        has_any = true;
                    }
                    b"left" => {
                        borders.left = border;
                        has_any = true;
                    }
                    b"right" => {
                        borders.right = border;
                        has_any = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"pBdr" => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }
    if has_any { Ok(Some(borders)) } else { Ok(None) }
}
