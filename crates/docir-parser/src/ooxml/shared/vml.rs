use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{local_name, xml_error};
use docir_core::ir::{VmlDrawing, VmlShape};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;
use quick_xml::Reader;

pub fn parse_vml_drawing(
    xml: &str,
    path: &str,
    rels: &Relationships,
) -> Result<(VmlDrawing, Vec<VmlShape>), ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    reader.config_mut().check_end_names = false;

    let mut drawing = VmlDrawing::new(path);
    drawing.span = Some(SourceSpan::new(path));
    let mut shapes: Vec<VmlShape> = Vec::new();
    let mut current: Option<VmlShape> = None;

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"shape" {
                    let mut shape = VmlShape::new();
                    shape.span = Some(SourceSpan::new(path));
                    apply_shape_attrs(&mut shape, &e);
                    current = Some(shape);
                } else if local == b"imagedata" {
                    if let Some(shape) = current.as_mut() {
                        apply_imagedata_attrs(shape, &e, rels);
                    }
                } else if local == b"textbox" {
                    if let Some(shape) = current.as_mut() {
                        let text = read_textbox_text(&mut reader)?;
                        if !text.is_empty() {
                            shape.text = Some(text);
                        }
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"shape" {
                    let mut shape = VmlShape::new();
                    shape.span = Some(SourceSpan::new(path));
                    apply_shape_attrs(&mut shape, &e);
                    shapes.push(shape);
                } else if local == b"imagedata" {
                    if let Some(shape) = current.as_mut() {
                        apply_imagedata_attrs(shape, &e, rels);
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"shape" {
                    if let Some(shape) = current.take() {
                        shapes.push(shape);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok((drawing, shapes))
}

fn apply_shape_attrs(shape: &mut VmlShape, e: &quick_xml::events::BytesStart<'_>) {
    for attr in e.attributes().flatten() {
        let key = local_name(attr.key.as_ref());
        let val = String::from_utf8_lossy(&attr.value).to_string();
        match key {
            b"id" | b"name" => shape.name = Some(val),
            b"type" => shape.shape_type = Some(val),
            b"style" => shape.style = Some(val),
            b"filled" => shape.filled = Some(val == "t" || val == "true" || val == "1"),
            b"stroked" => shape.stroked = Some(val == "t" || val == "true" || val == "1"),
            _ => {}
        }
    }
}

fn apply_imagedata_attrs(
    shape: &mut VmlShape,
    e: &quick_xml::events::BytesStart<'_>,
    rels: &Relationships,
) {
    for attr in e.attributes().flatten() {
        let key = local_name(attr.key.as_ref());
        let val = String::from_utf8_lossy(&attr.value).to_string();
        if key == b"id" || key == b"rid" || key == b"rId" {
            shape.rel_id = Some(val.clone());
            if let Some(rel) = rels.get(&val) {
                shape.image_target = Some(rel.target.clone());
            }
        }
    }
}

fn read_textbox_text(reader: &mut Reader<&[u8]>) -> Result<String, ParseError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Text(t)) => {
                text.push_str(&t.unescape().unwrap_or_default());
            }
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"t" {
                    if let Ok(Event::Text(t)) = reader.read_event_into(&mut buf) {
                        text.push_str(&t.unescape().unwrap_or_default());
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                if local_name(&name_buf) == b"textbox" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("vml_textbox", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}
