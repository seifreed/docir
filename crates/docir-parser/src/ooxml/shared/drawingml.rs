use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
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
    let mut current_diagram_rel_ids: Vec<String> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                match local {
                    b"docPr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"name" {
                                current_name =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    b"blip" => {
                        let mut rel_id = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"embed" || key == b"link" || key == b"id" {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        if let Some(rel_id) = rel_id {
                            if let Some(rel) = rels.get(&rel_id) {
                                let mut shape = Shape::new(ShapeType::Picture);
                                shape.name = current_name.clone();
                                shape.relationship_id = Some(rel_id);
                                shape.media_target =
                                    Some(resolve_drawingml_target(path, &rel.target));
                                shape.span = Some(SourceSpan::new(path));
                                shapes.push(shape);
                            }
                        }
                    }
                    b"relIds" => {
                        current_diagram_rel_ids.clear();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"dm" || key == b"lo" || key == b"qs" || key == b"cs" {
                                current_diagram_rel_ids
                                    .push(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        if !current_diagram_rel_ids.is_empty() {
                            let mut related_targets = Vec::new();
                            for rel_id in &current_diagram_rel_ids {
                                if let Some(rel) = rels.get(rel_id) {
                                    related_targets
                                        .push(resolve_drawingml_target(path, &rel.target));
                                }
                            }
                            let mut shape = Shape::new(ShapeType::Custom);
                            shape.name = current_name.clone();
                            shape.relationship_id = current_diagram_rel_ids.first().cloned();
                            shape.related_targets = related_targets;
                            shape.span = Some(SourceSpan::new(path));
                            shapes.push(shape);
                        }
                        current_diagram_rel_ids.clear();
                    }
                    b"chart" => {
                        let mut rel_id = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"id" || key == b"rid" {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        if let Some(rel_id) = rel_id {
                            if let Some(rel) = rels.get(&rel_id) {
                                let mut shape = Shape::new(ShapeType::Chart);
                                shape.name = current_name.clone();
                                shape.relationship_id = Some(rel_id);
                                shape.media_target =
                                    Some(resolve_drawingml_target(path, &rel.target));
                                shape.span = Some(SourceSpan::new(path));
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

fn normalize_docx_target(target: &str) -> String {
    let mut t = target;
    while t.starts_with("../") {
        t = &t[3..];
    }
    if t.starts_with("./") {
        t = &t[2..];
    }
    if t.starts_with("word/") {
        t.to_string()
    } else {
        format!("word/{}", t.trim_start_matches('/'))
    }
}
