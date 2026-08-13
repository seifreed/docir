use super::{
    DocxParser, NodeId, PageBorders, ParseError, Relationships, WordSettings,
    normalize_docx_target, parse_border, reader_from_str, span_from_reader,
};
use crate::xml_utils::local_name;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::try_attr_value;
use crate::xml_utils::try_attr_value_by_suffix;
use crate::xml_utils::visit_attributes;
use crate::xml_utils::xml_error;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

const DOC_PATH: &str = "word/document.xml";

pub(super) fn parse_page_borders(
    reader: &mut Reader<&[u8]>,
) -> Result<Option<PageBorders>, ParseError> {
    let mut buf = Vec::new();
    let mut borders = PageBorders::default();
    let mut has_any = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let border = parse_border(&e, DOC_PATH)?;
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
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"pgBorders" => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    if has_any { Ok(Some(borders)) } else { Ok(None) }
}

pub(super) fn bool_from_val(start: &BytesStart, file: &str) -> Result<bool, ParseError> {
    Ok(!matches!(
        try_attr_value(start, b"w:val", file)?.as_deref(),
        Some("0") | Some("false")
    ))
}

pub(super) fn parse_vml_style_length(style: &str, key: &str) -> Result<Option<i64>, ParseError> {
    Ok(parse_vml_style_length_value(style, key)?.map(|val| val.round() as i64))
}

fn parse_vml_style_length_u64(style: &str, key: &str) -> Result<Option<u64>, ParseError> {
    Ok(parse_vml_style_length_value(style, key)?
        .and_then(|val| (val >= 0.0).then(|| val.round() as u64)))
}

fn parse_vml_style_length_value(style: &str, key: &str) -> Result<Option<f64>, ParseError> {
    for part in style.split(';') {
        let mut iter = part.splitn(2, ':');
        let Some(style_key) = iter.next() else {
            continue;
        };
        let Some(style_value) = iter.next() else {
            continue;
        };
        let style_key = style_key.trim();
        let style_value = style_value.trim();
        if style_key.eq_ignore_ascii_case(key) {
            return parse_vml_length(style_value);
        }
    }
    Ok(None)
}

fn parse_vml_length(value: &str) -> Result<Option<f64>, ParseError> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return Ok(None);
    }
    let mut split_idx = trimmed.len();
    for (idx, ch) in trimmed.char_indices().rev() {
        if ch.is_ascii_alphabetic() {
            split_idx = idx;
        } else {
            break;
        }
    }
    let (num_part, unit) = if split_idx < trimmed.len() {
        trimmed.split_at(split_idx)
    } else {
        (trimmed, "")
    };
    let numeric_value = num_part.trim().parse::<f64>().map_err(|err| {
        ParseError::InvalidStructure(format!("Invalid VML length '{trimmed}': {err}"))
    })?;
    if !numeric_value.is_finite() {
        return Err(ParseError::InvalidStructure(format!(
            "Invalid non-finite VML length '{trimmed}'"
        )));
    }
    let unit = unit.trim();
    let emus = match unit {
        "" => numeric_value,
        "pt" => numeric_value * 12700.0,
        "in" => numeric_value * 914400.0,
        "cm" => numeric_value * 360000.0,
        "mm" => numeric_value * 36000.0,
        _ => {
            return Err(ParseError::InvalidStructure(format!(
                "Invalid VML length unit '{unit}' in '{trimmed}'"
            )));
        }
    };
    Ok(Some(emus))
}

pub(super) fn parse_vml_pict(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<Option<NodeId>, ParseError> {
    let mut buf = Vec::new();
    let mut rel_id: Option<String> = None;
    let mut name: Option<String> = None;
    let mut alt_text: Option<String> = None;
    let mut transform = docir_core::ir::ShapeTransform::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if local_name(e.name().as_ref()) == b"imagedata" {
                    rel_id = try_attr_value_by_suffix(&e, &[b":id"], "word/document.xml")?;
                } else if local_name(e.name().as_ref()) == b"shape" {
                    name = try_attr_value(&e, b"name", "word/document.xml")?.or(try_attr_value(
                        &e,
                        b"id",
                        "word/document.xml",
                    )?);
                    alt_text = try_attr_value(&e, b"o:title", "word/document.xml")?
                        .or(try_attr_value(&e, b"alt", "word/document.xml")?);
                    if let Some(style) = try_attr_value(&e, b"style", "word/document.xml")? {
                        if let Some(val) = parse_vml_style_length(&style, "left")? {
                            transform.x = val;
                        }
                        if let Some(val) = parse_vml_style_length(&style, "top")? {
                            transform.y = val;
                        }
                        if let Some(val) = parse_vml_style_length_u64(&style, "width")? {
                            transform.width = val;
                        }
                        if let Some(val) = parse_vml_style_length_u64(&style, "height")? {
                            transform.height = val;
                        }
                    }
                }
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"pict" => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    if let Some(rel_id) = rel_id
        && let Some(rel) = rels.get(&rel_id)
    {
        let mut shape = docir_core::ir::Shape::new(docir_core::ir::ShapeType::Picture);
        shape.name = name;
        shape.alt_text = alt_text;
        shape.transform = transform;
        shape.relationship_id = Some(rel_id.clone());
        shape.media_target = Some(normalize_docx_target(&rel.target));
        let mut span = span_from_reader(reader, "word/document.xml");
        span.relationship_id = Some(rel_id.clone());
        shape.span = Some(span);
        let shape_id = shape.id;
        parser.store.insert(docir_core::ir::IRNode::Shape(shape));
        return Ok(Some(shape_id));
    }
    Ok(None)
}

pub(super) fn parse_settings_like(xml: &str) -> Result<WordSettings, ParseError> {
    let mut settings = WordSettings::new();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                let mut entry = docir_core::ir::SettingEntry {
                    name,
                    value: None,
                    attributes: Vec::new(),
                };
                visit_attributes(&e, "word/settings.xml", |attr| {
                    let attr_name = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                    let attr_val = lossy_attr_value(attr).to_string();
                    entry.attributes.push(docir_core::ir::SettingAttribute {
                        name: attr_name,
                        value: attr_val,
                    });
                })?;
                settings.entries.push(entry);
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/settings.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(settings)
}

pub(super) fn parse_num_abstract_id(reader: &mut Reader<&[u8]>) -> Result<u32, ParseError> {
    let mut buf = Vec::new();
    let mut abstract_id = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                if local_name(e.name().as_ref()) == b"abstractNumId"
                    && let Some(val) = try_attr_value(&e, b"w:val", "word/numbering.xml")?
                        .map(|value| {
                            value
                                .parse::<u32>()
                                .map_err(|err| xml_error("word/numbering.xml", err))
                        })
                        .transpose()?
                {
                    abstract_id = Some(val);
                }
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"num" => {
                return abstract_id.ok_or_else(|| {
                    xml_error("word/numbering.xml", "num is missing w:abstractNumId")
                });
            }
            Ok(Event::Eof) => {
                return Err(xml_error(
                    "word/numbering.xml",
                    "unexpected end of numbering num",
                ));
            }
            Err(e) => {
                return Err(xml_error("word/numbering.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
}

pub(super) fn line_col(data: &[u8], pos: usize) -> Option<(u32, u32)> {
    if pos > data.len() {
        return None;
    }
    let slice = &data[..pos];
    let mut line = 1u32;
    let mut col = 1u32;
    for &b in slice {
        if b == b'\n' {
            line += 1;
            col = 1;
        } else {
            col = col.saturating_add(1);
        }
    }
    Some((line, col))
}
