use super::super::super::{DocxParser, Field, ParagraphProperties, span_from_reader};
use super::{DOC_XML_PATH, parse_field_instruction, parse_run};
use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::ooxml::relationships::TargetMode;
use crate::xml_utils::{
    attr_u32_from_bytes, local_name, try_attr_value, try_attr_value_by_suffix, xml_error,
};
use docir_core::ir::RunProperties;
use docir_core::ir::{Hyperlink, NumberingInfo, UnderlineStyle, VerticalTextAlignment};
use docir_core::types::NodeId;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

pub(crate) fn parse_numbering(
    reader: &mut Reader<&[u8]>,
    props: &mut ParagraphProperties,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    let mut num_id = None;
    let mut level = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"numId" => {
                    num_id =
                        try_attr_value(&e, b"w:val", DOC_XML_PATH)?.and_then(|v| v.parse().ok());
                }
                b"ilvl" => {
                    level =
                        try_attr_value(&e, b"w:val", DOC_XML_PATH)?.and_then(|v| v.parse().ok());
                }
                _ => {}
            },
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"numPr" => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_XML_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }

    if let (Some(num_id), Some(level)) = (num_id, level) {
        props.numbering = Some(NumberingInfo {
            num_id,
            level,
            format: None,
        });
    }
    Ok(())
}

pub(crate) fn parse_run_properties(
    reader: &mut Reader<&[u8]>,
    props: &mut RunProperties,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                b"rStyle" => {
                    if let Some(val) = try_attr_value(&e, b"w:val", DOC_XML_PATH)? {
                        props.style_id = Some(val);
                    }
                }
                b"rFonts" => {
                    if let Some(val) = try_attr_value(&e, b"w:ascii", DOC_XML_PATH)? {
                        props.font_family = Some(val);
                    }
                }
                b"sz" => {
                    if let Some(val) = attr_u32_from_bytes(&e, b"w:val", DOC_XML_PATH)? {
                        props.font_size = Some(val);
                    }
                }
                b"b" => props.bold = Some(true),
                b"i" => props.italic = Some(true),
                b"u" => {
                    if let Some(val) = try_attr_value(&e, b"w:val", DOC_XML_PATH)? {
                        props.underline = match val.as_str() {
                            "double" => Some(UnderlineStyle::Double),
                            "thick" => Some(UnderlineStyle::Thick),
                            "dotted" => Some(UnderlineStyle::Dotted),
                            "dash" | "dashed" => Some(UnderlineStyle::Dashed),
                            _ => Some(UnderlineStyle::Single),
                        };
                    } else {
                        props.underline = Some(UnderlineStyle::Single);
                    }
                }
                b"strike" => props.strike = Some(true),
                b"color" => {
                    if let Some(val) = try_attr_value(&e, b"w:val", DOC_XML_PATH)? {
                        props.color = Some(val);
                    }
                }
                b"highlight" => {
                    if let Some(val) = try_attr_value(&e, b"w:val", DOC_XML_PATH)? {
                        props.highlight = Some(val);
                    }
                }
                b"caps" => {
                    props.all_caps = Some(try_bool_from_val(&e)?);
                }
                b"smallCaps" => {
                    props.small_caps = Some(try_bool_from_val(&e)?);
                }
                b"vertAlign" => {
                    if let Some(val) = try_attr_value(&e, b"w:val", DOC_XML_PATH)? {
                        props.vertical_align = match val.as_str() {
                            "superscript" => Some(VerticalTextAlignment::Superscript),
                            "subscript" => Some(VerticalTextAlignment::Subscript),
                            _ => None,
                        };
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"rPr" => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_XML_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn try_bool_from_val(start: &BytesStart<'_>) -> Result<bool, ParseError> {
    Ok(!matches!(
        try_attr_value(start, b"w:val", DOC_XML_PATH)?.as_deref(),
        Some("0") | Some("false")
    ))
}

pub(crate) fn parse_hyperlink(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart,
) -> Result<NodeId, ParseError> {
    let mut link = Hyperlink::new("", false);
    let mut rel_id_opt = None;
    if let Some(tooltip) = try_attr_value(start, b"w:tooltip", DOC_XML_PATH)? {
        link.tooltip = Some(tooltip);
    }
    if let Some(rel_id) = try_attr_value_by_suffix(start, &[b":id"], DOC_XML_PATH)?
        && let Some(rel) = rels.get(&rel_id)
    {
        link.target = rel.target.clone();
        link.is_external = rel.target_mode == TargetMode::External;
        link.relationship_id = Some(rel_id.clone());
        rel_id_opt = Some(rel_id);
    }
    if let Some(anchor) = try_attr_value(start, b"w:anchor", DOC_XML_PATH)? {
        if link.target.is_empty() {
            link.target = format!("#{}", anchor);
        } else if !link.target.contains('#') {
            link.target = format!("{}#{}", link.target, anchor);
        }
    }
    let mut span = span_from_reader(reader, DOC_XML_PATH);
    span.relationship_id = rel_id_opt.clone();
    link.span = Some(span);

    let mut buf = Vec::new();
    super::scan_docx_xml_events_until_end_start_only(
        reader,
        &mut buf,
        |event| matches!(event, Event::End(e) if local_name(e.name().as_ref()) == b"hyperlink"),
        |reader, start| {
            if local_name(start.name().as_ref()) == b"r" {
                let run = parse_run(parser, reader, rels)?;
                link.runs.push(run.run_id);
            }
            Ok(())
        },
        |_reader, _event| Ok(()),
    )?;

    let id = link.id;
    parser.store.insert(docir_core::ir::IRNode::Hyperlink(link));
    Ok(id)
}

pub(crate) fn parse_field(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    instruction: Option<String>,
) -> Result<NodeId, ParseError> {
    let mut field = Field::new(instruction);
    field.instruction_parsed = field
        .instruction
        .as_deref()
        .and_then(parse_field_instruction);
    let mut buf = Vec::new();

    super::scan_docx_xml_events_until_end_start_only(
        reader,
        &mut buf,
        |event| matches!(event, Event::End(e) if local_name(e.name().as_ref()) == b"fldSimple"),
        |reader, start| {
            if local_name(start.name().as_ref()) == b"r" {
                let run = parse_run(parser, reader, &Relationships::default())?;
                field.runs.push(run.run_id);
            }
            Ok(())
        },
        |_reader, _event| Ok(()),
    )?;

    let id = field.id;
    parser.store.insert(docir_core::ir::IRNode::Field(field));
    Ok(id)
}
