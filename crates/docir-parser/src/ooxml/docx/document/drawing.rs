use super::{normalize_docx_target, span_from_reader, DocxParser};
use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{attr_value, xml_error};
use docir_core::ir::{
    Shape, ShapeText, ShapeTextParagraph, ShapeTextRun, ShapeTransform, ShapeType, TextAlignment,
};
use docir_core::types::NodeId;
use quick_xml::events::Event;
use quick_xml::Reader;

pub(super) fn parse_drawing(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<Option<NodeId>, ParseError> {
    let mut buf = Vec::new();
    let mut rel_id: Option<String> = None;
    let mut chart_rel: Option<String> = None;
    let mut diagram_rel_ids: Vec<String> = Vec::new();
    let mut name: Option<String> = None;
    let mut alt_text: Option<String> = None;
    let mut shape_type = ShapeType::Picture;
    let mut transform = ShapeTransform::default();
    let mut next_pos_is_x = true;
    let mut text: Option<ShapeText> = None;
    let mut hyperlink_rel: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_bytes = e.name().as_ref().to_vec();
                let n = name_bytes.as_slice();
                if n == b"a:blip" {
                    rel_id = attr_value(&e, b"r:embed").or_else(|| attr_value(&e, b"r:link"));
                } else if n == b"wp:docPr" {
                    name = attr_value(&e, b"name");
                    alt_text = attr_value(&e, b"descr");
                } else if n == b"a:graphicData" {
                    if let Some(uri) = attr_value(&e, b"uri") {
                        if uri.contains("chart") {
                            shape_type = docir_core::ir::ShapeType::Chart;
                        } else if uri.contains("diagram") {
                            shape_type = docir_core::ir::ShapeType::Custom;
                        }
                    }
                } else if n == b"a:prstGeom" {
                    if let Some(val) = attr_value(&e, b"prst") {
                        shape_type = map_shape_type(&val);
                    }
                } else if n == b"wp:extent" || n == b"a:ext" {
                    if let Some(val) = attr_value(&e, b"cx").and_then(|v| v.parse().ok()) {
                        transform.width = val;
                    }
                    if let Some(val) = attr_value(&e, b"cy").and_then(|v| v.parse().ok()) {
                        transform.height = val;
                    }
                } else if n == b"a:off" {
                    if let Some(val) = attr_value(&e, b"x").and_then(|v| v.parse().ok()) {
                        transform.x = val;
                    }
                    if let Some(val) = attr_value(&e, b"y").and_then(|v| v.parse().ok()) {
                        transform.y = val;
                    }
                } else if n == b"wp:posOffset" {
                    if let Ok(text) = reader.read_text(e.name()) {
                        if let Ok(val) = text.parse::<i64>() {
                            if next_pos_is_x {
                                transform.x = val;
                            } else {
                                transform.y = val;
                            }
                            next_pos_is_x = !next_pos_is_x;
                        }
                    }
                } else if n == b"a:txBody" {
                    text = Some(parse_drawing_text_body(reader, "word/document.xml")?);
                } else if n.ends_with(b":chart") || n == b"c:chart" {
                    chart_rel = attr_value(&e, b"r:id");
                } else if n == b"dgm:relIds" {
                    if let Some(val) = attr_value(&e, b"r:dm") {
                        diagram_rel_ids.push(val);
                    }
                    if let Some(val) = attr_value(&e, b"r:lo") {
                        diagram_rel_ids.push(val);
                    }
                    if let Some(val) = attr_value(&e, b"r:qs") {
                        diagram_rel_ids.push(val);
                    }
                    if let Some(val) = attr_value(&e, b"r:cs") {
                        diagram_rel_ids.push(val);
                    }
                } else if n == b"a:hlinkClick" {
                    hyperlink_rel = attr_value(&e, b"r:id");
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:drawing" {
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

    let rel_id = chart_rel
        .clone()
        .or(diagram_rel_ids.first().cloned())
        .or(rel_id);
    if let Some(rel_id) = rel_id {
        if let Some(rel) = rels.get(&rel_id) {
            let mut shape = Shape::new(shape_type);
            shape.name = name;
            shape.alt_text = alt_text;
            shape.transform = transform;
            shape.text = text;
            shape.relationship_id = Some(rel_id.clone());
            shape.media_target = Some(normalize_docx_target(&rel.target));
            let mut span = span_from_reader(reader, "word/document.xml");
            span.relationship_id = Some(rel_id.clone());
            shape.span = Some(span);
            if let Some(hrel) = hyperlink_rel.as_ref().and_then(|id| rels.get(id)) {
                shape.hyperlink = Some(hrel.target.clone());
            }
            if !diagram_rel_ids.is_empty() {
                let mut related_targets = Vec::new();
                for rel_id in diagram_rel_ids {
                    if let Some(rel) = rels.get(&rel_id) {
                        related_targets.push(normalize_docx_target(&rel.target));
                    }
                }
                shape.related_targets = related_targets;
            }
            let shape_id = shape.id;
            parser.store.insert(docir_core::ir::IRNode::Shape(shape));
            return Ok(Some(shape_id));
        }
    }
    Ok(None)
}

fn parse_drawing_text_body(
    reader: &mut Reader<&[u8]>,
    doc_path: &str,
) -> Result<ShapeText, ParseError> {
    let mut paragraphs = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"a:p" {
                    let paragraph = parse_drawing_text_paragraph(reader, doc_path)?;
                    paragraphs.push(paragraph);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"a:txBody" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(doc_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(ShapeText { paragraphs })
}

fn parse_drawing_text_paragraph(
    reader: &mut Reader<&[u8]>,
    doc_path: &str,
) -> Result<ShapeTextParagraph, ParseError> {
    let mut runs = Vec::new();
    let mut alignment: Option<TextAlignment> = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"a:pPr" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"algn" {
                            alignment = map_alignment(&String::from_utf8_lossy(&attr.value));
                        }
                    }
                }
                b"a:r" => {
                    let run = parse_drawing_text_run(reader, doc_path)?;
                    runs.push(run);
                }
                b"a:br" => {
                    runs.push(ShapeTextRun {
                        text: "\n".to_string(),
                        bold: None,
                        italic: None,
                        font_size: None,
                        font_family: None,
                    });
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"a:p" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(doc_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(ShapeTextParagraph { runs, alignment })
}

fn parse_drawing_text_run(
    reader: &mut Reader<&[u8]>,
    doc_path: &str,
) -> Result<ShapeTextRun, ParseError> {
    let mut text = String::new();
    let mut bold = None;
    let mut italic = None;
    let mut font_size = None;
    let mut font_family = None;

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"a:t" => {
                    let t = reader.read_text(e.name()).map_err(|e| ParseError::Xml {
                        file: doc_path.to_string(),
                        message: e.to_string(),
                    })?;
                    text.push_str(&t);
                }
                b"a:rPr" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"b" => bold = Some(attr.value.as_ref() == b"1"),
                            b"i" => italic = Some(attr.value.as_ref() == b"1"),
                            b"sz" => {
                                font_size = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            _ => {}
                        }
                    }
                }
                b"a:latin" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"typeface" {
                            font_family = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"a:r" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(doc_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(ShapeTextRun {
        text,
        bold,
        italic,
        font_size,
        font_family,
    })
}

fn map_shape_type(value: &str) -> ShapeType {
    match value {
        "rect" => ShapeType::Rectangle,
        "roundRect" => ShapeType::RoundRect,
        "ellipse" => ShapeType::Ellipse,
        "triangle" => ShapeType::Triangle,
        "line" => ShapeType::Line,
        "straightConnector1" => ShapeType::Line,
        "bentConnector2" | "bentConnector3" | "bentConnector4" | "bentConnector5" => {
            ShapeType::Line
        }
        "rightArrow" | "leftArrow" | "upArrow" | "downArrow" | "leftRightArrow" | "upDownArrow"
        | "bentArrow" | "uTurnArrow" | "curvedRightArrow" | "curvedLeftArrow" | "curvedUpArrow"
        | "curvedDownArrow" => ShapeType::Arrow,
        _ => ShapeType::Custom,
    }
}

fn map_alignment(value: &str) -> Option<TextAlignment> {
    match value {
        "l" => Some(TextAlignment::Left),
        "ctr" => Some(TextAlignment::Center),
        "r" => Some(TextAlignment::Right),
        "just" => Some(TextAlignment::Justify),
        "dist" => Some(TextAlignment::Distribute),
        _ => None,
    }
}
