use super::{span_from_reader, DocxParser};
use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::ooxml::shared::normalize_docx_target;
use crate::xml_utils::{attr_bool_like, attr_u32_from_bytes, attr_value, xml_error};
use docir_core::ir::{
    Shape, ShapeText, ShapeTextParagraph, ShapeTextRun, ShapeTransform, ShapeType, TextAlignment,
};
use docir_core::types::NodeId;
use quick_xml::events::{BytesStart, Event};
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
                let name_slice = name_bytes.as_slice();
                if name_slice == b"a:blip" {
                    rel_id = attr_value(&e, b"r:embed").or_else(|| attr_value(&e, b"r:link"));
                } else if name_slice == b"wp:docPr" {
                    name = attr_value(&e, b"name");
                    alt_text = attr_value(&e, b"descr");
                } else if name_slice == b"a:graphicData" {
                    if let Some(uri) = attr_value(&e, b"uri") {
                        if uri.contains("chart") {
                            shape_type = docir_core::ir::ShapeType::Chart;
                        } else if uri.contains("diagram") {
                            shape_type = docir_core::ir::ShapeType::Custom;
                        }
                    }
                } else if name_slice == b"a:prstGeom" {
                    if let Some(val) = attr_value(&e, b"prst") {
                        shape_type = map_shape_type(&val);
                    }
                } else if name_slice == b"wp:extent" || name_slice == b"a:ext" {
                    if let Some(val) = attr_value(&e, b"cx").and_then(|v| v.parse().ok()) {
                        transform.width = val;
                    }
                    if let Some(val) = attr_value(&e, b"cy").and_then(|v| v.parse().ok()) {
                        transform.height = val;
                    }
                } else if name_slice == b"a:off" {
                    if let Some(val) = attr_value(&e, b"x").and_then(|v| v.parse().ok()) {
                        transform.x = val;
                    }
                    if let Some(val) = attr_value(&e, b"y").and_then(|v| v.parse().ok()) {
                        transform.y = val;
                    }
                } else if name_slice == b"wp:posOffset" {
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
                } else if name_slice == b"a:txBody" {
                    text = Some(parse_drawing_text_body(reader, "word/document.xml")?);
                } else if name_slice.ends_with(b":chart") || name_slice == b"c:chart" {
                    chart_rel = attr_value(&e, b"r:id");
                } else if name_slice == b"dgm:relIds" {
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
                } else if name_slice == b"a:hlinkClick" {
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
                    alignment = parse_paragraph_alignment(&e);
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
                    let t = reader
                        .read_text(e.name())
                        .map_err(|e| xml_error(doc_path, e))?;
                    text.push_str(&t);
                }
                b"a:rPr" => {
                    parse_run_style_attrs(&e, &mut bold, &mut italic, &mut font_size);
                }
                b"a:latin" => {
                    font_family = parse_run_font_family(&e);
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

fn parse_paragraph_alignment(start: &BytesStart<'_>) -> Option<TextAlignment> {
    attr_value(start, b"algn").and_then(|value| map_alignment(&value))
}

fn parse_run_style_attrs(
    start: &BytesStart<'_>,
    bold: &mut Option<bool>,
    italic: &mut Option<bool>,
    font_size: &mut Option<u32>,
) {
    if let Some(value) = attr_value(start, b"b") {
        *bold = Some(attr_bool_like(value.as_bytes()));
    }
    if let Some(value) = attr_value(start, b"i") {
        *italic = Some(attr_bool_like(value.as_bytes()));
    }
    *font_size = attr_u32_from_bytes(start, b"sz");
}

fn parse_run_font_family(start: &BytesStart<'_>) -> Option<String> {
    attr_value(start, b"typeface")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml_utils::reader_from_str;

    #[test]
    fn map_shape_type_covers_connectors_and_arrow_variants() {
        assert_eq!(map_shape_type("straightConnector1"), ShapeType::Line);
        assert_eq!(map_shape_type("bentConnector3"), ShapeType::Line);
        assert_eq!(map_shape_type("rightArrow"), ShapeType::Arrow);
        assert_eq!(map_shape_type("curvedUpArrow"), ShapeType::Arrow);
        assert_eq!(map_shape_type("rect"), ShapeType::Rectangle);
        assert_eq!(map_shape_type("unknownShape"), ShapeType::Custom);
    }

    #[test]
    fn map_alignment_maps_known_values_and_unknown_to_none() {
        assert_eq!(map_alignment("l"), Some(TextAlignment::Left));
        assert_eq!(map_alignment("ctr"), Some(TextAlignment::Center));
        assert_eq!(map_alignment("r"), Some(TextAlignment::Right));
        assert_eq!(map_alignment("just"), Some(TextAlignment::Justify));
        assert_eq!(map_alignment("dist"), Some(TextAlignment::Distribute));
        assert_eq!(map_alignment("x"), None);
    }

    #[test]
    fn parse_drawing_text_run_parses_text_and_run_style_flags() {
        let xml = r#"
            <a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:rPr b="0" i="1" sz="bad"></a:rPr>
              <a:latin typeface="Calibri"></a:latin>
              <a:t>Hello</a:t>
            </a:r>
        "#;
        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"a:r" => break,
                Ok(Event::Eof) => panic!("a:r start not found"),
                Ok(_) => {}
                Err(err) => panic!("unexpected xml read error: {err}"),
            }
            buf.clear();
        }

        let run =
            parse_drawing_text_run(&mut reader, "word/document.xml").expect("drawing run parse");
        assert_eq!(run.text, "Hello");
        assert_eq!(run.bold, Some(false));
        assert_eq!(run.italic, Some(true));
        assert_eq!(run.font_size, None);
        assert_eq!(run.font_family.as_deref(), Some("Calibri"));
    }

    #[test]
    fn parse_drawing_text_body_parses_alignment_runs_and_breaks() {
        let xml = r#"
            <a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:p>
                <a:pPr algn="ctr"></a:pPr>
                <a:r><a:t>Line1</a:t></a:r>
                <a:br></a:br>
                <a:r><a:t>Line2</a:t></a:r>
              </a:p>
            </a:txBody>
        "#;
        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"a:txBody" => break,
                Ok(Event::Eof) => panic!("a:txBody start not found"),
                Ok(_) => {}
                Err(err) => panic!("unexpected xml read error: {err}"),
            }
            buf.clear();
        }

        let text =
            parse_drawing_text_body(&mut reader, "word/document.xml").expect("text body parse");
        assert_eq!(text.paragraphs.len(), 1);
        assert_eq!(text.paragraphs[0].alignment, Some(TextAlignment::Center));
        assert_eq!(text.paragraphs[0].runs.len(), 3);
        assert_eq!(text.paragraphs[0].runs[0].text, "Line1");
        assert_eq!(text.paragraphs[0].runs[1].text, "\n");
        assert_eq!(text.paragraphs[0].runs[2].text, "Line2");
    }
}
