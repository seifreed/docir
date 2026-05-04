use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::{local_name, xml_error};
use docir_core::ir::{VmlDrawing, VmlShape};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;
use quick_xml::Reader;

/// Public API entrypoint: parse_vml_drawing.
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
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                handle_vml_element_start(&mut current, local, &e, path, rels, &mut reader, false)?;
            }
            Ok(Event::Empty(e)) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if let Some(shape) = handle_vml_element_start(
                    &mut current,
                    local,
                    &e,
                    path,
                    rels,
                    &mut reader,
                    true,
                )? {
                    shapes.push(shape);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
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

fn handle_vml_element_start(
    current: &mut Option<VmlShape>,
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
    path: &str,
    rels: &Relationships,
    reader: &mut Reader<&[u8]>,
    is_empty: bool,
) -> Result<Option<VmlShape>, ParseError> {
    match local {
        b"shape" => {
            let mut shape = VmlShape::new();
            shape.span = Some(SourceSpan::new(path));
            apply_shape_attrs(&mut shape, e);
            if is_empty {
                return Ok(Some(shape));
            }
            *current = Some(shape);
            Ok(None)
        }
        b"imagedata" => {
            if let Some(shape) = current.as_mut() {
                apply_imagedata_attrs(shape, e, rels);
            }
            Ok(None)
        }
        b"textbox" => {
            if let Some(shape) = current.as_mut() {
                let text = read_textbox_text(reader)?;
                if !text.is_empty() {
                    shape.text = Some(text);
                }
            }
            Ok(None)
        }
        _ => Ok(None),
    }
}

fn apply_shape_attrs(shape: &mut VmlShape, e: &quick_xml::events::BytesStart<'_>) {
    for attr in e.attributes().flatten() {
        let key = local_name(attr.key.as_ref());
        let val = lossy_attr_value(&attr).to_string();
        match key {
            b"id" | b"name" => shape.name = Some(val),
            b"type" => shape.shape_type = Some(val),
            b"style" => shape.style = Some(val),
            b"filled" => shape.filled = Some(parse_shape_bool_attr(&val)),
            b"stroked" => shape.stroked = Some(parse_shape_bool_attr(&val)),
            _ => {}
        }
    }
}

fn parse_shape_bool_attr(value: &str) -> bool {
    value == "t" || value == "true" || value == "1"
}

fn apply_imagedata_attrs(
    shape: &mut VmlShape,
    e: &quick_xml::events::BytesStart<'_>,
    rels: &Relationships,
) {
    let Some(rel_id) = parse_image_rel_id(e) else {
        return;
    };
    shape.rel_id = Some(rel_id.clone());
    if let Some(rel) = rels.get(&rel_id) {
        shape.image_target = Some(rel.target.clone());
    }
}

fn parse_image_rel_id(e: &quick_xml::events::BytesStart<'_>) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = local_name(attr.key.as_ref());
        if key != b"id" && key != b"rid" && key != b"rId" {
            continue;
        }
        return Some(lossy_attr_value(&attr).to_string());
    }
    None
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::relationships::{Relationship, TargetMode};
    use std::collections::HashMap;

    fn relationships_with_image() -> Relationships {
        let mut by_id = HashMap::new();
        by_id.insert(
            "rId5".to_string(),
            Relationship {
                id: "rId5".to_string(),
                rel_type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                        .to_string(),
                target: "media/image1.png".to_string(),
                target_mode: TargetMode::Internal,
            },
        );
        Relationships {
            by_id,
            by_type: HashMap::new(),
        }
    }

    #[test]
    fn parse_vml_drawing_extracts_shape_style_and_image_target() {
        let xml = r##"
            <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <v:shape id="shape1" type="#_x0000_t75" style="position:absolute" filled="t" stroked="0">
                <v:imagedata r:id="rId5"/>
              </v:shape>
            </xml>
        "##;

        let (drawing, shapes) =
            parse_vml_drawing(xml, "word/vmlDrawing1.vml", &relationships_with_image())
                .expect("vml");

        assert_eq!(drawing.path, "word/vmlDrawing1.vml");
        assert_eq!(shapes.len(), 1);
        let shape = &shapes[0];
        assert_eq!(shape.name.as_deref(), Some("shape1"));
        assert_eq!(shape.shape_type.as_deref(), Some("#_x0000_t75"));
        assert_eq!(shape.style.as_deref(), Some("position:absolute"));
        assert_eq!(shape.filled, Some(true));
        assert_eq!(shape.stroked, Some(false));
        assert_eq!(shape.rel_id.as_deref(), Some("rId5"));
        assert_eq!(shape.image_target.as_deref(), Some("media/image1.png"));
    }

    #[test]
    fn parse_vml_drawing_extracts_textbox_text() {
        let xml = r#"
            <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word">
              <v:shape id="shape2">
                <v:textbox>
                  <w:txbxContent><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:txbxContent>
                </v:textbox>
              </v:shape>
            </xml>
        "#;

        let (_, shapes) =
            parse_vml_drawing(xml, "word/vmlDrawing2.vml", &Relationships::default()).expect("vml");

        assert_eq!(shapes.len(), 1);
        assert_eq!(shapes[0].text.as_deref(), Some("Hello"));
    }

    #[test]
    fn parse_vml_drawing_is_tolerant_for_incomplete_xml() {
        let (drawing, shapes) =
            parse_vml_drawing("<v:shape>", "word/bad.vml", &Relationships::default())
                .expect("parser is best-effort");
        assert_eq!(drawing.path, "word/bad.vml");
        assert!(shapes.is_empty());
    }
}
