use super::*;

pub(super) fn parse_page_borders(
    reader: &mut Reader<&[u8]>,
) -> Result<Option<PageBorders>, ParseError> {
    let mut buf = Vec::new();
    let mut borders = PageBorders::default();
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
                if e.name().as_ref() == b"w:pgBorders" {
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
    if has_any {
        Ok(Some(borders))
    } else {
        Ok(None)
    }
}

pub(super) fn bool_from_val(start: &BytesStart) -> bool {
    match attr_value(start, b"w:val").as_deref() {
        Some("0") | Some("false") => false,
        _ => true,
    }
}

pub(super) fn parse_vml_style_length(style: &str, key: &str) -> Option<i64> {
    parse_vml_style_length_value(style, key).map(|val| val.round() as i64)
}

fn parse_vml_style_length_u64(style: &str, key: &str) -> Option<u64> {
    parse_vml_style_length_value(style, key).and_then(|val| {
        if val >= 0.0 {
            Some(val.round() as u64)
        } else {
            None
        }
    })
}

fn parse_vml_style_length_value(style: &str, key: &str) -> Option<f64> {
    for part in style.split(';') {
        let mut iter = part.splitn(2, ':');
        let k = iter.next()?.trim();
        let v = iter.next()?.trim();
        if k.eq_ignore_ascii_case(key) {
            return parse_vml_length(v);
        }
    }
    None
}

fn parse_vml_length(value: &str) -> Option<f64> {
    let v = value.trim();
    if v.is_empty() {
        return None;
    }
    let mut split_idx = v.len();
    for (idx, ch) in v.char_indices().rev() {
        if ch.is_ascii_alphabetic() {
            split_idx = idx;
        } else {
            break;
        }
    }
    let (num_part, unit) = if split_idx < v.len() {
        v.split_at(split_idx)
    } else {
        (v, "")
    };
    let value = num_part.trim().parse::<f64>().ok()?;
    let unit = unit.trim();
    let emus = match unit {
        "" => value,
        "pt" => value * 12700.0,
        "in" => value * 914400.0,
        "cm" => value * 360000.0,
        "mm" => value * 36000.0,
        _ => return None,
    };
    Some(emus)
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
                if e.name().as_ref() == b"v:imagedata" {
                    rel_id = attr_value(&e, b"r:id");
                } else if e.name().as_ref() == b"v:shape" {
                    name = attr_value(&e, b"name").or_else(|| attr_value(&e, b"id"));
                    alt_text = attr_value(&e, b"o:title").or_else(|| attr_value(&e, b"alt"));
                    if let Some(style) = attr_value(&e, b"style") {
                        if let Some(val) = parse_vml_style_length(&style, "left") {
                            transform.x = val;
                        }
                        if let Some(val) = parse_vml_style_length(&style, "top") {
                            transform.y = val;
                        }
                        if let Some(val) = parse_vml_style_length_u64(&style, "width") {
                            transform.width = val;
                        }
                        if let Some(val) = parse_vml_style_length_u64(&style, "height") {
                            transform.height = val;
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:pict" {
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

    if let Some(rel_id) = rel_id {
        if let Some(rel) = rels.get(&rel_id) {
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
                for attr in e.attributes().flatten() {
                    let attr_name = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                    let attr_val = String::from_utf8_lossy(&attr.value).to_string();
                    entry.attributes.push(docir_core::ir::SettingAttribute {
                        name: attr_name,
                        value: attr_val,
                    });
                }
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
    let mut abstract_id = 0u32;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"w:abstractNumId" {
                    if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                        abstract_id = val;
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:num" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/numbering.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(abstract_id)
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
