use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::{local_name, visit_attributes, xml_error};
use docir_core::ir::{VmlDrawing, VmlShape};
use docir_core::types::SourceSpan;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Public API entrypoint: parse_vml_drawing.
pub fn parse_vml_drawing(
    xml: &str,
    path: &str,
    rels: &Relationships,
) -> Result<(VmlDrawing, Vec<VmlShape>), ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut drawing = VmlDrawing::new(path);
    drawing.span = Some(SourceSpan::new(path));
    let mut shapes: Vec<VmlShape> = Vec::new();
    let mut current: Option<VmlShape> = None;

    let mut buf = Vec::new();
    let mut root_name: Option<Vec<u8>> = None;
    let mut root_closed = false;
    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error(path, err))?;
        if root_closed && matches!(&event, Event::Start(_) | Event::Empty(_)) {
            return Err(xml_error(path, "XML document contains multiple roots"));
        }
        match &event {
            Event::Start(e) if root_name.is_none() => {
                root_name = Some(e.name().as_ref().to_vec());
            }
            Event::Empty(e) if root_name.is_none() => {
                root_name = Some(e.name().as_ref().to_vec());
                root_closed = true;
            }
            Event::End(e)
                if root_name
                    .as_deref()
                    .is_some_and(|name| name == e.name().as_ref()) =>
            {
                root_closed = true;
            }
            Event::Eof if root_closed => break,
            Event::Eof => {
                return Err(xml_error(
                    path,
                    "XML document ends before its root is closed",
                ));
            }
            _ => {}
        }
        match event {
            Event::Start(e) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                handle_vml_element_start(&mut current, local, &e, path, rels, &mut reader, false)?;
            }
            Event::Empty(e) => {
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
            Event::End(e) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if local == b"shape"
                    && let Some(shape) = current.take()
                {
                    shapes.push(shape);
                }
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
            apply_shape_attrs(&mut shape, e, path)?;
            if is_empty {
                return Ok(Some(shape));
            }
            *current = Some(shape);
            Ok(None)
        }
        b"imagedata" => {
            if let Some(shape) = current.as_mut() {
                apply_imagedata_attrs(shape, e, rels, path)?;
            }
            Ok(None)
        }
        b"textbox" => {
            if let Some(shape) = current.as_mut() {
                let text = read_textbox_text(reader, path)?;
                if !text.is_empty() {
                    shape.text = Some(text);
                }
            }
            Ok(None)
        }
        _ => Ok(None),
    }
}

fn apply_shape_attrs(
    shape: &mut VmlShape,
    e: &quick_xml::events::BytesStart<'_>,
    path: &str,
) -> Result<(), ParseError> {
    visit_attributes(e, path, |attr| {
        let key = local_name(attr.key.as_ref());
        let val = lossy_attr_value(attr).to_string();
        match key {
            b"id" | b"name" => shape.name = Some(val),
            b"type" => shape.shape_type = Some(val),
            b"style" => shape.style = Some(val),
            b"filled" => shape.filled = Some(parse_shape_bool_attr(&val)),
            b"stroked" => shape.stroked = Some(parse_shape_bool_attr(&val)),
            _ => {}
        }
    })?;
    Ok(())
}

fn parse_shape_bool_attr(value: &str) -> bool {
    value == "t" || value == "true" || value == "1"
}

fn apply_imagedata_attrs(
    shape: &mut VmlShape,
    e: &quick_xml::events::BytesStart<'_>,
    rels: &Relationships,
    path: &str,
) -> Result<(), ParseError> {
    let Some(rel_id) = parse_image_rel_id(e, path)? else {
        return Ok(());
    };
    shape.rel_id = Some(rel_id.clone());
    if let Some(rel) = rels.get(&rel_id) {
        shape.image_target = Some(rel.target.clone());
    }
    Ok(())
}

fn parse_image_rel_id(
    e: &quick_xml::events::BytesStart<'_>,
    path: &str,
) -> Result<Option<String>, ParseError> {
    let mut rel_id = None;
    visit_attributes(e, path, |attr| {
        let key = local_name(attr.key.as_ref());
        if rel_id.is_none() && matches!(key, b"id" | b"rid" | b"rId") {
            rel_id = Some(lossy_attr_value(attr).to_string());
        }
    })?;
    Ok(rel_id)
}

fn read_textbox_text(reader: &mut Reader<&[u8]>, path: &str) -> Result<String, ParseError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Text(t)) => {
                text.push_str(
                    &crate::xml_utils::decoded_text(&t).map_err(|err| xml_error(path, err))?,
                );
            }
            Ok(Event::GeneralRef(t)) => {
                text.push_str(
                    &crate::xml_utils::decoded_general_ref(&t)
                        .map_err(|err| xml_error(path, err))?,
                );
            }
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"t" {
                    match reader.read_event_into(&mut buf) {
                        Ok(Event::Text(t)) => {
                            text.push_str(
                                &crate::xml_utils::decoded_text(&t)
                                    .map_err(|err| xml_error(path, err))?,
                            );
                        }
                        Ok(Event::GeneralRef(t)) => {
                            text.push_str(
                                &crate::xml_utils::decoded_general_ref(&t)
                                    .map_err(|err| xml_error(path, err))?,
                            );
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                if local_name(&name_buf) == b"textbox" {
                    break;
                }
            }
            Ok(Event::Eof) => {
                return Err(xml_error(path, "unexpected end of textbox"));
            }
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
    fn parse_vml_drawing_rejects_incomplete_xml() {
        let err = parse_vml_drawing("<v:shape>", "word/bad.vml", &Relationships::default())
            .expect_err("truncated VML must fail");
        assert!(matches!(err, ParseError::Xml { file, .. } if file == "word/bad.vml"));

        let err = parse_vml_drawing(
            "<xml/><xml/>",
            "word/multiple-roots.vml",
            &Relationships::default(),
        )
        .expect_err("multiple VML roots must fail");
        assert!(matches!(err, ParseError::Xml { file, .. } if file == "word/multiple-roots.vml"));
    }

    #[test]
    fn parse_vml_drawing_reports_malformed_attributes() {
        let err = parse_vml_drawing(
            r#"<xml><v:shape id="shape1" id="shape2"/></xml>"#,
            "word/vmlDrawing1.vml",
            &Relationships::default(),
        )
        .expect_err("malformed vml shape attrs should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "word/vmlDrawing1.vml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }
}
