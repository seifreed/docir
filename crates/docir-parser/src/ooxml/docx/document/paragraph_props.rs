use super::super::{apply_section_refs, bool_from_val, parse_border, parse_numbering, SectionRef};
use crate::error::ParseError;
use crate::xml_utils::attr_value;
use docir_core::ir::Paragraph;
use docir_core::ir::{LineSpacingRule, TextAlignment};
use docir_core::types::NodeId;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;

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
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:pStyle" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        para.style_id = Some(val);
                    }
                }
                b"w:keepNext" => {
                    para.properties.keep_next = Some(bool_from_val(&e));
                }
                b"w:keepLines" => {
                    para.properties.keep_lines = Some(bool_from_val(&e));
                }
                b"w:pageBreakBefore" => {
                    para.properties.page_break_before = Some(bool_from_val(&e));
                }
                b"w:widowControl" => {
                    para.properties.widow_control = Some(bool_from_val(&e));
                }
                b"w:jc" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        para.properties.alignment = Some(alignment_from_val(val.as_str()));
                    }
                }
                b"w:ind" => {
                    let mut indent = para.properties.indentation.clone().unwrap_or_default();
                    if let Some(val) = attr_value(&e, b"w:left").and_then(|v| v.parse().ok()) {
                        indent.left = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:right").and_then(|v| v.parse().ok()) {
                        indent.right = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:firstLine").and_then(|v| v.parse().ok()) {
                        indent.first_line = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:hanging").and_then(|v| v.parse().ok()) {
                        indent.hanging = Some(val);
                    }
                    para.properties.indentation = Some(indent);
                }
                b"w:spacing" => {
                    let mut spacing = para.properties.spacing.clone().unwrap_or_default();
                    if let Some(val) = attr_value(&e, b"w:before").and_then(|v| v.parse().ok()) {
                        spacing.before = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:after").and_then(|v| v.parse().ok()) {
                        spacing.after = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:line").and_then(|v| v.parse().ok()) {
                        spacing.line = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:lineRule") {
                        spacing.line_rule = match val.as_str() {
                            "auto" => Some(LineSpacingRule::Auto),
                            "exact" => Some(LineSpacingRule::Exact),
                            "atLeast" => Some(LineSpacingRule::AtLeast),
                            _ => None,
                        };
                    }
                    para.properties.spacing = Some(spacing);
                }
                b"w:pBdr" => {
                    if let Some(borders) = parse_paragraph_borders(reader)? {
                        para.properties.borders = Some(borders);
                    }
                }
                b"w:outlineLvl" => {
                    if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                        para.properties.outline_level = Some(val);
                    }
                }
                b"w:numPr" => {
                    parse_numbering(reader, &mut para.properties)?;
                }
                b"w:sectPr" => {
                    section_ref = Some(apply_section_refs(reader, header_footer_map)?);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:pPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(section_ref)
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
                match e.name().as_ref() {
                    b"w:top" => {
                        borders.top = border;
                        has_any = true;
                    }
                    b"w:bottom" => {
                        borders.bottom = border;
                        has_any = true;
                    }
                    b"w:left" => {
                        borders.left = border;
                        has_any = true;
                    }
                    b"w:right" => {
                        borders.right = border;
                        has_any = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:pBdr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    if has_any {
        Ok(Some(borders))
    } else {
        Ok(None)
    }
}
