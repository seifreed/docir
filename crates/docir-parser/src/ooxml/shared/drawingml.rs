use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::ooxml::shared::normalize_docx_target;
use crate::xml_utils::{local_name, xml_error};
use docir_core::ir::{DrawingPart, Shape, ShapeType};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;
use quick_xml::Reader;

pub fn parse_drawingml_part(
    xml: &str,
    path: &str,
    rels: &Relationships,
) -> Result<(DrawingPart, Vec<Shape>), ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut part = DrawingPart::new(path);
    part.span = Some(SourceSpan::new(path));
    let mut shapes: Vec<Shape> = Vec::new();

    let mut buf = Vec::new();
    let mut current_name: Option<String> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                match local {
                    b"docPr" => {
                        current_name = attr_value_by_local_keys(&e, &[b"name"]);
                    }
                    b"blip" => {
                        if let Some(rel_id) =
                            attr_value_by_local_keys(&e, &[b"embed", b"link", b"id"])
                        {
                            if let Some(shape) = build_target_shape(
                                rels,
                                path,
                                current_name.clone(),
                                rel_id,
                                ShapeType::Picture,
                            ) {
                                shapes.push(shape);
                            }
                        }
                    }
                    b"relIds" => {
                        let rel_ids = rel_ids_from_attr_keys(&e, &[b"dm", b"lo", b"qs", b"cs"]);
                        if !rel_ids.is_empty() {
                            shapes.push(build_custom_shape(
                                rels,
                                path,
                                current_name.clone(),
                                rel_ids,
                            ));
                        }
                    }
                    b"chart" => {
                        if let Some(rel_id) = attr_value_by_local_keys(&e, &[b"id", b"rid"]) {
                            if let Some(shape) = build_target_shape(
                                rels,
                                path,
                                current_name.clone(),
                                rel_id,
                                ShapeType::Chart,
                            ) {
                                shapes.push(shape);
                            }
                        }
                    }
                    _ => {}
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

    Ok((part, shapes))
}

fn resolve_drawingml_target(path: &str, target: &str) -> String {
    if path.starts_with("word/") {
        normalize_docx_target(target)
    } else {
        Relationships::resolve_target(path, target)
    }
}

fn attr_value_by_local_keys(
    event: &quick_xml::events::BytesStart<'_>,
    keys: &[&[u8]],
) -> Option<String> {
    for attr in event.attributes().flatten() {
        let key = local_name(attr.key.as_ref());
        if keys.iter().any(|candidate| key == *candidate) {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

fn rel_ids_from_attr_keys(
    event: &quick_xml::events::BytesStart<'_>,
    keys: &[&[u8]],
) -> Vec<String> {
    event
        .attributes()
        .flatten()
        .filter_map(|attr| {
            let key = local_name(attr.key.as_ref());
            if keys.iter().any(|candidate| key == *candidate) {
                Some(String::from_utf8_lossy(&attr.value).to_string())
            } else {
                None
            }
        })
        .collect()
}

fn build_target_shape(
    rels: &Relationships,
    path: &str,
    current_name: Option<String>,
    rel_id: String,
    shape_type: ShapeType,
) -> Option<Shape> {
    let rel = rels.get(&rel_id)?;
    let mut shape = Shape::new(shape_type);
    shape.name = current_name;
    shape.relationship_id = Some(rel_id);
    shape.media_target = Some(resolve_drawingml_target(path, &rel.target));
    shape.span = Some(SourceSpan::new(path));
    Some(shape)
}

fn build_custom_shape(
    rels: &Relationships,
    path: &str,
    current_name: Option<String>,
    rel_ids: Vec<String>,
) -> Shape {
    let related_targets = rel_ids
        .iter()
        .filter_map(|rel_id| rels.get(rel_id))
        .map(|rel| resolve_drawingml_target(path, &rel.target))
        .collect();
    let mut shape = Shape::new(ShapeType::Custom);
    shape.name = current_name;
    shape.relationship_id = rel_ids.first().cloned();
    shape.related_targets = related_targets;
    shape.span = Some(SourceSpan::new(path));
    shape
}
