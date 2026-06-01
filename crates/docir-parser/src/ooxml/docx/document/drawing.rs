use super::{DocxParser, span_from_reader};
use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::ooxml::shared::normalize_docx_target;
use crate::xml_utils::{
    attr_bool_like, attr_u32_from_bytes, local_name, try_attr_value, try_attr_value_by_suffix,
    xml_error,
};
use docir_core::ir::{
    Shape, ShapeText, ShapeTextParagraph, ShapeTextRun, ShapeTransform, ShapeType, TextAlignment,
};
use docir_core::types::NodeId;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

const DOC_PATH: &str = "word/document.xml";

pub(super) fn parse_drawing(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<Option<NodeId>, ParseError> {
    let mut buf = Vec::new();
    let mut state = DocxDrawingState::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                handle_drawing_start(reader, &e, &mut state)?;
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"drawing" => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }

    finish_drawing_shape(parser, reader, rels, state)
}

struct DocxDrawingState {
    rel_id: Option<String>,
    chart_rel: Option<String>,
    diagram_rel_ids: Vec<String>,
    name: Option<String>,
    alt_text: Option<String>,
    shape_type: ShapeType,
    transform: ShapeTransform,
    next_pos_is_x: bool,
    text: Option<ShapeText>,
    hyperlink_rel: Option<String>,
}

impl DocxDrawingState {
    fn new() -> Self {
        Self {
            rel_id: None,
            chart_rel: None,
            diagram_rel_ids: Vec::new(),
            name: None,
            alt_text: None,
            shape_type: ShapeType::Picture,
            transform: ShapeTransform::default(),
            next_pos_is_x: true,
            text: None,
            hyperlink_rel: None,
        }
    }
}

fn handle_drawing_start(
    reader: &mut Reader<&[u8]>,
    e: &BytesStart<'_>,
    state: &mut DocxDrawingState,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"blip" => state.rel_id = try_attr_value_by_suffix(e, &[b":embed", b":link"], DOC_PATH)?,
        b"docPr" => {
            state.name = try_attr_value(e, b"name", DOC_PATH)?;
            state.alt_text = try_attr_value(e, b"descr", DOC_PATH)?;
        }
        b"graphicData" => apply_graphic_data_type(e, state)?,
        b"prstGeom" => {
            if let Some(val) = try_attr_value(e, b"prst", DOC_PATH)? {
                state.shape_type = map_shape_type(&val);
            }
        }
        b"extent" | b"ext" => apply_extent(e, &mut state.transform)?,
        b"off" => apply_offset(e, &mut state.transform)?,
        b"posOffset" => apply_position_offset(reader, e, state),
        b"txBody" => state.text = Some(parse_drawing_text_body(reader, DOC_PATH)?),
        b"chart" => state.chart_rel = try_attr_value_by_suffix(e, &[b":id"], DOC_PATH)?,
        b"relIds" => collect_diagram_relationships(e, &mut state.diagram_rel_ids)?,
        b"hlinkClick" => state.hyperlink_rel = try_attr_value_by_suffix(e, &[b":id"], DOC_PATH)?,
        _ => {}
    }
    Ok(())
}

fn apply_graphic_data_type(
    e: &BytesStart<'_>,
    state: &mut DocxDrawingState,
) -> Result<(), ParseError> {
    if let Some(uri) = try_attr_value(e, b"uri", DOC_PATH)? {
        if uri.contains("chart") {
            state.shape_type = ShapeType::Chart;
        } else if uri.contains("diagram") {
            state.shape_type = ShapeType::Custom;
        }
    }
    Ok(())
}

fn apply_extent(e: &BytesStart<'_>, transform: &mut ShapeTransform) -> Result<(), ParseError> {
    if let Some(val) = try_attr_value(e, b"cx", DOC_PATH)?.and_then(|v| v.parse().ok()) {
        transform.width = val;
    }
    if let Some(val) = try_attr_value(e, b"cy", DOC_PATH)?.and_then(|v| v.parse().ok()) {
        transform.height = val;
    }
    Ok(())
}

fn apply_offset(e: &BytesStart<'_>, transform: &mut ShapeTransform) -> Result<(), ParseError> {
    if let Some(val) = try_attr_value(e, b"x", DOC_PATH)?.and_then(|v| v.parse().ok()) {
        transform.x = val;
    }
    if let Some(val) = try_attr_value(e, b"y", DOC_PATH)?.and_then(|v| v.parse().ok()) {
        transform.y = val;
    }
    Ok(())
}

fn apply_position_offset(
    reader: &mut Reader<&[u8]>,
    e: &BytesStart<'_>,
    state: &mut DocxDrawingState,
) {
    if let Ok(text) = reader.read_text(e.name())
        && let Ok(val) = text.parse::<i64>()
    {
        if state.next_pos_is_x {
            state.transform.x = val;
        } else {
            state.transform.y = val;
        }
        state.next_pos_is_x = !state.next_pos_is_x;
    }
}

fn collect_diagram_relationships(
    e: &BytesStart<'_>,
    diagram_rel_ids: &mut Vec<String>,
) -> Result<(), ParseError> {
    for suffix in [
        b":dm".as_slice(),
        b":lo".as_slice(),
        b":qs".as_slice(),
        b":cs".as_slice(),
    ] {
        if let Some(val) = try_attr_value_by_suffix(e, &[suffix], DOC_PATH)? {
            diagram_rel_ids.push(val);
        }
    }
    Ok(())
}

fn finish_drawing_shape(
    parser: &mut DocxParser,
    reader: &Reader<&[u8]>,
    rels: &Relationships,
    state: DocxDrawingState,
) -> Result<Option<NodeId>, ParseError> {
    let rel_id = state
        .chart_rel
        .clone()
        .or(state.diagram_rel_ids.first().cloned())
        .or(state.rel_id);
    if let Some(rel_id) = rel_id
        && let Some(rel) = rels.get(&rel_id)
    {
        let mut shape = Shape::new(state.shape_type);
        shape.name = state.name;
        shape.alt_text = state.alt_text;
        shape.transform = state.transform;
        shape.text = state.text;
        shape.relationship_id = Some(rel_id.clone());
        shape.media_target = Some(normalize_docx_target(&rel.target));
        let mut span = span_from_reader(reader, DOC_PATH);
        span.relationship_id = Some(rel_id.clone());
        shape.span = Some(span);
        if let Some(hrel) = state.hyperlink_rel.as_ref().and_then(|id| rels.get(id)) {
            shape.hyperlink = Some(hrel.target.clone());
        }
        if !state.diagram_rel_ids.is_empty() {
            let mut related_targets = Vec::new();
            for rel_id in state.diagram_rel_ids {
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
            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"p" => {
                let paragraph = parse_drawing_text_paragraph(reader, doc_path)?;
                paragraphs.push(paragraph);
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"txBody" => {
                break;
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
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"pPr" => {
                    alignment = parse_paragraph_alignment(&e, doc_path)?;
                }
                b"r" => {
                    let run = parse_drawing_text_run(reader, doc_path)?;
                    runs.push(run);
                }
                b"br" => {
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
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"p" => {
                break;
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
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"t" => {
                    let t = reader
                        .read_text(e.name())
                        .map_err(|e| xml_error(doc_path, e))?;
                    text.push_str(&t);
                }
                b"rPr" => {
                    parse_run_style_attrs(&e, doc_path, &mut bold, &mut italic, &mut font_size)?;
                }
                b"latin" => {
                    font_family = parse_run_font_family(&e, doc_path)?;
                }
                _ => {}
            },
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"r" => {
                break;
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

fn parse_paragraph_alignment(
    start: &BytesStart<'_>,
    doc_path: &str,
) -> Result<Option<TextAlignment>, ParseError> {
    Ok(try_attr_value(start, b"algn", doc_path)?.and_then(|value| map_alignment(&value)))
}

fn parse_run_style_attrs(
    start: &BytesStart<'_>,
    doc_path: &str,
    bold: &mut Option<bool>,
    italic: &mut Option<bool>,
    font_size: &mut Option<u32>,
) -> Result<(), ParseError> {
    if let Some(value) = try_attr_value(start, b"b", doc_path)? {
        *bold = Some(attr_bool_like(value.as_bytes()));
    }
    if let Some(value) = try_attr_value(start, b"i", doc_path)? {
        *italic = Some(attr_bool_like(value.as_bytes()));
    }
    *font_size = attr_u32_from_bytes(start, b"sz", doc_path)?;
    Ok(())
}

fn parse_run_font_family(
    start: &BytesStart<'_>,
    doc_path: &str,
) -> Result<Option<String>, ParseError> {
    try_attr_value(start, b"typeface", doc_path)
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
