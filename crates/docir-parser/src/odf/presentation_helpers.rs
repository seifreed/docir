//! ODF presentation parsing helpers extracted from the main module.

use super::helpers::{parse_notes, parse_text_element};
use crate::xml_utils::{attr_value_by_suffix, local_name, xml_error};
#[path = "presentation_helpers_utils.rs"]
mod presentation_helpers_utils;
use super::{
    ChartData, IRNode, IrStore, MediaAsset, MediaType, NodeId, OdfReader, ParseError, Shape,
    ShapeText, ShapeTextParagraph, ShapeTextRun, ShapeType, Slide, SlideAnimation, SlideTransition,
    SourceSpan, parse_frame_transform, read_event,
};
#[path = "presentation_helpers_frame.rs"]
mod presentation_helpers_frame;
pub(crate) use presentation_helpers_frame::classify_media_shape;
pub(crate) use presentation_helpers_frame::{
    FrameShapeState, parse_frame_shape_empty, parse_frame_shape_start,
};
use presentation_helpers_utils::{classify_media_type, parse_duration_ms};
use quick_xml::events::{BytesStart, Event};

pub(super) fn parse_draw_page(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    slide_no: u32,
    store: &mut IrStore,
) -> Result<Slide, ParseError> {
    let mut slide = Slide::new(slide_no);
    slide.name = attr_value_by_suffix(start, &[b":name"]);
    slide.master_id = attr_value_by_suffix(start, &[b":master-page-name"]);
    slide.layout_id = attr_value_by_suffix(start, &[b":page-layout-name"])
        .or_else(|| attr_value_by_suffix(start, &[b":style-name"]));
    slide.transition = parse_odp_transition(start);

    let mut state = DrawPageState {
        slide,
        notes_text: None,
    };
    let mut buf = Vec::new();

    loop {
        match read_event(reader, &mut buf, "content.xml")? {
            Event::Start(e) => handle_draw_page_start_event(reader, &e, store, &mut state)?,
            Event::Empty(e) => handle_draw_page_empty_event(reader, &e, store, &mut state)?,
            Event::End(e) if local_name(e.name().as_ref()) == b"page" => {
                break;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    state.slide.notes = state.notes_text;
    Ok(state.slide)
}

struct DrawPageState {
    slide: Slide,
    notes_text: Option<String>,
}

fn handle_draw_page_start_event(
    reader: &mut OdfReader<'_>,
    event: &BytesStart<'_>,
    store: &mut IrStore,
    state: &mut DrawPageState,
) -> Result<(), ParseError> {
    match local_name(event.name().as_ref()) {
        b"frame" => {
            if let Some(shape_id) = parse_draw_frame_presentation(reader, event, store)? {
                state.slide.shapes.push(shape_id);
            }
        }
        b"custom-shape" => {
            if let Some(shape_id) = parse_custom_shape_presentation(reader, event, store)? {
                state.slide.shapes.push(shape_id);
            }
        }
        b"notes" => {
            state.notes_text = parse_notes(reader)?;
        }
        name if name.starts_with(b"anim:") => {
            if let Some(anim) = parse_odf_animation(event) {
                state.slide.animations.push(anim);
            }
        }
        _ => {}
    }
    Ok(())
}

fn handle_draw_page_empty_event(
    reader: &mut OdfReader<'_>,
    event: &BytesStart<'_>,
    store: &mut IrStore,
    state: &mut DrawPageState,
) -> Result<(), ParseError> {
    match local_name(event.name().as_ref()) {
        b"frame" => {
            if let Some(shape_id) = parse_draw_frame_presentation(reader, event, store)? {
                state.slide.shapes.push(shape_id);
            }
        }
        b"custom-shape" => {
            if let Some(shape_id) = parse_custom_shape_presentation(reader, event, store)? {
                state.slide.shapes.push(shape_id);
            }
        }
        name if name.starts_with(b"anim:") => {
            if let Some(anim) = parse_odf_animation(event) {
                state.slide.animations.push(anim);
            }
        }
        _ => {}
    }
    Ok(())
}

pub(super) fn parse_draw_frame_presentation(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let transform = parse_frame_transform(start);
    let mut state = DrawFrameState {
        frame: FrameShapeState::new(),
        text: None,
        name: attr_value_by_suffix(start, &[b":name"]),
    };
    let mut buf = Vec::new();

    loop {
        match read_event(reader, &mut buf, "content.xml")? {
            Event::Start(e) => handle_draw_frame_start_event(reader, &e, store, &mut state)?,
            Event::Empty(e) => parse_frame_shape_empty(&e, store, &mut state.frame),
            Event::End(e) if local_name(e.name().as_ref()) == b"frame" => {
                break;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if state.frame.has_shape {
        let mut shape = Shape::new(state.frame.shape_type);
        shape.name = state.name;
        shape.media_target = state.frame.media_target;
        shape.text = state.text;
        shape.chart_id = state.frame.chart_id;
        shape.transform = transform;
        let id = shape.id;
        store.insert(IRNode::Shape(shape));
        Ok(Some(id))
    } else {
        Ok(None)
    }
}

struct DrawFrameState {
    frame: FrameShapeState,
    text: Option<ShapeText>,
    name: Option<String>,
}

fn handle_draw_frame_start_event(
    reader: &mut OdfReader<'_>,
    event: &BytesStart<'_>,
    store: &mut IrStore,
    state: &mut DrawFrameState,
) -> Result<(), ParseError> {
    match local_name(event.name().as_ref()) {
        b"text-box" => {
            let paragraphs = parse_shape_text(reader, event.name().as_ref())?;
            if !paragraphs.is_empty() {
                state.text = Some(ShapeText { paragraphs });
                state.frame.shape_type = ShapeType::TextBox;
                state.frame.has_shape = true;
            }
        }
        _ => parse_frame_shape_start(reader, event, store, &mut state.frame)?,
    }
    Ok(())
}

pub(super) fn parse_custom_shape_presentation(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let mut name = attr_value_by_suffix(start, &[b":name"]);
    let paragraphs = parse_shape_text(reader, start.name().as_ref())?;
    let mut shape = Shape::new(ShapeType::Custom);
    shape.name = name.take();
    if !paragraphs.is_empty() {
        shape.text = Some(ShapeText { paragraphs });
    }
    let shape_id = shape.id;
    store.insert(IRNode::Shape(shape));
    Ok(Some(shape_id))
}

fn parse_shape_text(
    reader: &mut OdfReader<'_>,
    end_tag: &[u8],
) -> Result<Vec<ShapeTextParagraph>, ParseError> {
    let mut buf = Vec::new();
    let mut paragraphs = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"p" => {
                let text = parse_text_element(reader, e.name().as_ref())?;
                let run = ShapeTextRun {
                    text,
                    bold: None,
                    italic: None,
                    font_size: None,
                    font_family: None,
                };
                paragraphs.push(ShapeTextParagraph {
                    runs: vec![run],
                    alignment: None,
                });
            }
            Ok(Event::End(e)) if e.name().as_ref() == end_tag => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error("content.xml", e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(paragraphs)
}

pub(super) fn parse_odp_transition(start: &BytesStart<'_>) -> Option<SlideTransition> {
    let transition_type = attr_value_by_suffix(start, &[b":transition-type"]);
    let speed = attr_value_by_suffix(start, &[b":transition-speed"]);
    let duration_ms =
        attr_value_by_suffix(start, &[b":transition-duration"]).and_then(|v| v.parse::<u32>().ok());
    let advance_after_ms =
        attr_value_by_suffix(start, &[b":duration"]).and_then(|v| v.parse::<u32>().ok());
    let advance_on_click = attr_value_by_suffix(start, &[b":animation"]).map(|v| v == "click");
    if transition_type.is_some() || speed.is_some() || duration_ms.is_some() {
        Some(SlideTransition {
            transition_type,
            speed,
            advance_on_click,
            advance_after_ms,
            duration_ms,
        })
    } else {
        None
    }
}

pub(super) fn parse_odf_chart(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
) -> Result<ChartData, ParseError> {
    let mut chart = ChartData::new();
    chart.chart_type = attr_value_by_suffix(start, &[b":class"]);
    chart.span = Some(SourceSpan::new("content.xml"));
    let mut buf = Vec::new();
    let mut in_title = false;
    let mut title_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"title" => {
                    in_title = true;
                }
                b"p" if in_title => {
                    let text = parse_text_element(reader, e.name().as_ref())?;
                    if !title_text.is_empty() && !text.is_empty() {
                        title_text.push(' ');
                    }
                    title_text.push_str(&text);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"title" {
                    in_title = false;
                }
                if local_name(e.name().as_ref()) == b"chart" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error("content.xml", e)),
            _ => {}
        }
        buf.clear();
    }

    if !title_text.is_empty() {
        chart.title = Some(title_text);
    }
    Ok(chart)
}

pub(super) fn parse_odf_animation(start: &BytesStart<'_>) -> Option<SlideAnimation> {
    let name = String::from_utf8_lossy(start.name().as_ref()).to_string();
    let mut anim = SlideAnimation {
        animation_type: name,
        target: attr_value_by_suffix(start, &[b":targetElement"]),
        duration_ms: attr_value_by_suffix(start, &[b":dur"]).and_then(|v| parse_duration_ms(&v)),
        preset_id: attr_value_by_suffix(start, &[b":preset-id"]),
        preset_class: attr_value_by_suffix(start, &[b":preset-class"]),
        media_asset: None,
    };
    if anim.target.is_none() {
        anim.target = attr_value_by_suffix(start, &[b":targetElement"]);
    }
    Some(anim)
}

pub(super) fn build_media_asset(path: &str, media: &str, size_bytes: u64) -> Option<MediaAsset> {
    let media_type = classify_media_type(path, media)?;
    let mut asset = MediaAsset::new(path.to_string(), media_type, size_bytes);
    asset.content_type = Some(media.to_string());
    asset.span = Some(SourceSpan::new("META-INF/manifest.xml"));
    Some(asset)
}

#[cfg(test)]
#[path = "presentation_helpers_tests.rs"]
mod tests;
