//! ODF presentation parsing helpers extracted from the main module.

use super::helpers::{parse_notes, parse_text_element};
use crate::xml_utils::{attr_value_by_suffix, local_name, xml_error};
#[path = "presentation_helpers_utils.rs"]
mod presentation_helpers_utils;
use super::{
    parse_frame_transform, read_event, ChartData, IRNode, IrStore, MediaAsset, MediaType, NodeId,
    OdfReader, ParseError, Shape, ShapeText, ShapeTextParagraph, ShapeTextRun, ShapeType, Slide,
    SlideAnimation, SlideTransition, SourceSpan,
};
#[path = "presentation_helpers_frame.rs"]
mod presentation_helpers_frame;
pub(crate) use presentation_helpers_frame::classify_media_shape;
pub(crate) use presentation_helpers_frame::{
    parse_frame_shape_empty, parse_frame_shape_start, FrameShapeState,
};
use presentation_helpers_utils::{classify_media_type, parse_duration_ms};
use quick_xml::events::{BytesStart, Event};
#[cfg(test)]
use quick_xml::Reader;

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
            Event::End(e) => {
                if local_name(e.name().as_ref()) == b"page" {
                    break;
                }
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
            Event::End(e) => {
                if local_name(e.name().as_ref()) == b"frame" {
                    break;
                }
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
            Ok(Event::Start(e)) => {
                if local_name(e.name().as_ref()) == b"p" {
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
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_tag {
                    break;
                }
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
mod tests {
    use super::*;

    fn parse_page_start(xml: &[u8]) -> (Reader<std::io::Cursor<&[u8]>>, BytesStart<'static>) {
        let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let page_start = loop {
            match reader.read_event_into(&mut buf).unwrap() {
                Event::Start(e) if local_name(e.name().as_ref()) == b"page" => {
                    break e.into_owned();
                }
                Event::Eof => panic!("missing draw:page"),
                _ => {}
            }
            buf.clear();
        };
        (reader, page_start)
    }

    #[test]
    fn parse_draw_page_extracts_metadata_transition_notes_and_shape_text() {
        let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <d:page d:name="SlideA"
    p:master-page-name="MasterA"
    d:style-name="LayoutFallback"
    p:transition-type="fade"
    p:transition-duration="2"
    p:animation="click">
    <d:frame d:name="TitleShape">
      <d:text-box>
        <t:p>Hello ODP</t:p>
      </d:text-box>
    </d:frame>
    <p:notes>
      <t:p>Speaker notes</t:p>
    </p:notes>
  </d:page>
</office:document-content>
"#;

        let (mut reader, page_start) = parse_page_start(xml);
        let mut store = IrStore::new();

        let slide = parse_draw_page(&mut reader, &page_start, 3, &mut store).unwrap();
        assert_eq!(slide.number, 3);
        assert_eq!(slide.name.as_deref(), Some("SlideA"));
        assert_eq!(slide.master_id.as_deref(), Some("MasterA"));
        assert_eq!(slide.layout_id.as_deref(), Some("LayoutFallback"));
        assert_eq!(slide.notes.as_deref(), Some("Speaker notes"));
        assert_eq!(slide.shapes.len(), 1);

        let transition = slide.transition.expect("expected transition");
        assert_eq!(transition.transition_type.as_deref(), Some("fade"));
        assert_eq!(transition.duration_ms, Some(2));
        assert_eq!(transition.advance_on_click, Some(true));

        let Some(IRNode::Shape(shape)) = store.get(slide.shapes[0]) else {
            panic!("expected shape node");
        };
        assert_eq!(shape.name.as_deref(), Some("TitleShape"));
        assert_eq!(shape.shape_type, ShapeType::TextBox);
        let text = shape.text.as_ref().expect("expected shape text");
        assert_eq!(text.paragraphs.len(), 1);
        assert_eq!(text.paragraphs[0].runs[0].text, "Hello ODP");
    }

    #[test]
    fn parse_draw_frame_presentation_returns_none_for_unrecognized_content() {
        let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<draw:frame xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">
  <draw:unknown/>
</draw:frame>
"#;
        let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let frame_start = loop {
            match reader.read_event_into(&mut buf).unwrap() {
                Event::Start(e) if e.name().as_ref() == b"draw:frame" => break e.into_owned(),
                Event::Eof => panic!("missing draw:frame"),
                _ => {}
            }
            buf.clear();
        };
        let mut store = IrStore::new();

        let shape = parse_draw_frame_presentation(&mut reader, &frame_start, &mut store).unwrap();
        assert!(shape.is_none());
        assert_eq!(store.values().count(), 0);
    }

    #[test]
    fn parse_draw_frame_presentation_classifies_plugin_media() {
        let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<draw:frame xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  draw:name="ClipFrame">
  <draw:plugin xlink:href="media/clip.mp4"/>
</draw:frame>
"#;
        let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let frame_start = loop {
            match reader.read_event_into(&mut buf).unwrap() {
                Event::Start(e) if e.name().as_ref() == b"draw:frame" => break e.into_owned(),
                Event::Eof => panic!("missing draw:frame"),
                _ => {}
            }
            buf.clear();
        };
        let mut store = IrStore::new();

        let shape_id =
            parse_draw_frame_presentation(&mut reader, &frame_start, &mut store).unwrap();
        let Some(shape_id) = shape_id else {
            panic!("expected shape");
        };
        let Some(IRNode::Shape(shape)) = store.get(shape_id) else {
            panic!("expected shape node");
        };
        assert_eq!(shape.name.as_deref(), Some("ClipFrame"));
        assert_eq!(shape.media_target.as_deref(), Some("media/clip.mp4"));
        assert_eq!(shape.shape_type, ShapeType::Video);
    }

    #[test]
    fn parse_custom_shape_presentation_preserves_text_runs() {
        let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<draw:custom-shape xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  draw:name="Badge">
  <text:p>First line</text:p>
  <text:p>Second line</text:p>
</draw:custom-shape>
"#;
        let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let shape_start = loop {
            match reader.read_event_into(&mut buf).unwrap() {
                Event::Start(e) if e.name().as_ref() == b"draw:custom-shape" => {
                    break e.into_owned();
                }
                Event::Eof => panic!("missing draw:custom-shape"),
                _ => {}
            }
            buf.clear();
        };
        let mut store = IrStore::new();

        let shape_id =
            parse_custom_shape_presentation(&mut reader, &shape_start, &mut store).unwrap();
        let Some(shape_id) = shape_id else {
            panic!("expected custom shape");
        };
        let Some(IRNode::Shape(shape)) = store.get(shape_id) else {
            panic!("expected shape node");
        };
        assert_eq!(shape.name.as_deref(), Some("Badge"));
        assert_eq!(shape.shape_type, ShapeType::Custom);
        let text = shape.text.as_ref().expect("shape text");
        assert_eq!(text.paragraphs.len(), 2);
        assert_eq!(text.paragraphs[0].runs[0].text, "First line");
        assert_eq!(text.paragraphs[1].runs[0].text, "Second line");
    }

    #[test]
    fn classify_media_shape_covers_audio_video_and_unknown_paths() {
        assert_eq!(classify_media_shape("media/clip.ogg"), ShapeType::Audio);
        assert_eq!(classify_media_shape("media/clip.OGV"), ShapeType::Video);
        assert_eq!(classify_media_shape("media/blob.bin"), ShapeType::Unknown);
    }

    #[test]
    fn parse_odf_animation_prefers_target_fallback_and_parses_iso_duration() {
        let mut start = BytesStart::new("anim:animate");
        start.push_attribute(("smil:targetElement", "shape-42"));
        start.push_attribute(("smil:dur", "PT2.5S"));
        start.push_attribute(("presentation:preset-id", "entrance"));

        let anim = parse_odf_animation(&start).expect("animation metadata");
        assert_eq!(anim.target.as_deref(), Some("shape-42"));
        assert_eq!(anim.duration_ms, Some(2500));
        assert_eq!(anim.preset_id.as_deref(), Some("entrance"));
    }

    #[test]
    fn parse_odp_transition_supports_fallback_keys_and_ignores_advance_only() {
        let mut fallback = BytesStart::new("draw:page");
        fallback.push_attribute(("draw:transition-type", "wipe"));
        fallback.push_attribute(("presentation:transition-speed", "fast"));
        fallback.push_attribute(("presentation:transition-duration", "900"));
        fallback.push_attribute(("presentation:duration", "1200"));
        fallback.push_attribute(("presentation:animation", "click"));

        let transition = parse_odp_transition(&fallback).expect("transition");
        assert_eq!(transition.transition_type.as_deref(), Some("wipe"));
        assert_eq!(transition.speed.as_deref(), Some("fast"));
        assert_eq!(transition.duration_ms, Some(900));
        assert_eq!(transition.advance_after_ms, Some(1200));
        assert_eq!(transition.advance_on_click, Some(true));

        let mut advance_only = BytesStart::new("draw:page");
        advance_only.push_attribute(("presentation:duration", "1000"));
        assert!(parse_odp_transition(&advance_only).is_none());
    }

    #[test]
    fn parse_duration_and_media_classification_cover_helper_paths() {
        assert_eq!(parse_duration_ms("250ms"), Some(250));
        assert_eq!(parse_duration_ms("1.25s"), Some(1250));
        assert_eq!(parse_duration_ms("PT3S"), Some(3000));
        assert_eq!(parse_duration_ms(""), None);
        assert_eq!(parse_duration_ms("invalid"), None);

        assert_eq!(
            classify_media_type("media/clip.oga", "application/ogg"),
            Some(MediaType::Audio)
        );
        assert_eq!(
            classify_media_type("media/clip.ogv", "application/ogg"),
            Some(MediaType::Video)
        );
        assert_eq!(
            classify_media_type("media/blob.bin", "application/octet-stream"),
            None
        );

        let asset = build_media_asset("Pictures/p1.png", "image/png", 42).expect("asset");
        assert_eq!(asset.media_type, MediaType::Image);
        assert_eq!(asset.content_type.as_deref(), Some("image/png"));
        assert_eq!(
            asset.span.as_ref().map(|s| s.file_path.as_str()),
            Some("META-INF/manifest.xml")
        );
    }

    #[test]
    fn parse_frame_shape_empty_covers_image_chart_and_ole_variants() {
        let mut frame = FrameShapeState::new();
        let mut store = IrStore::new();

        let mut image = BytesStart::new("dr:image");
        image.push_attribute(("lnk:href", "Pictures/img1.png"));
        parse_frame_shape_empty(&image, &mut store, &mut frame);
        assert_eq!(frame.shape_type, ShapeType::Picture);
        assert_eq!(frame.media_target.as_deref(), Some("Pictures/img1.png"));
        assert!(frame.has_shape);

        let mut chart = BytesStart::new("ch:chart");
        chart.push_attribute(("ch:class", "bar"));
        parse_frame_shape_empty(&chart, &mut store, &mut frame);
        assert_eq!(frame.shape_type, ShapeType::Chart);
        let chart_id = frame.chart_id.expect("chart id");
        let Some(IRNode::ChartData(chart_node)) = store.get(chart_id) else {
            panic!("expected chart data node");
        };
        assert_eq!(chart_node.chart_type.as_deref(), Some("bar"));

        let object = BytesStart::new("dr:object-ole");
        parse_frame_shape_empty(&object, &mut store, &mut frame);
        assert_eq!(frame.shape_type, ShapeType::OleObject);
        assert!(frame.has_shape);
    }

    #[test]
    fn parse_frame_shape_start_extracts_chart_title_text() {
        let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<chart:chart xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  chart:class="chart:line">
  <chart:title>
    <text:p>Main</text:p>
    <text:p>Title</text:p>
  </chart:title>
</chart:chart>
"#;
        let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let start = loop {
            match reader.read_event_into(&mut buf).unwrap() {
                Event::Start(e) if e.name().as_ref() == b"chart:chart" => break e.into_owned(),
                Event::Eof => panic!("missing chart:chart"),
                _ => {}
            }
            buf.clear();
        };

        let mut frame = FrameShapeState::new();
        let mut store = IrStore::new();
        parse_frame_shape_start(&mut reader, &start, &mut store, &mut frame).expect("parse");

        assert_eq!(frame.shape_type, ShapeType::Chart);
        assert!(frame.has_shape);
        let chart_id = frame.chart_id.expect("chart id");
        let Some(IRNode::ChartData(chart)) = store.get(chart_id) else {
            panic!("expected chart data");
        };
        assert_eq!(chart.chart_type.as_deref(), Some("chart:line"));
        assert_eq!(chart.title.as_deref(), Some("Main Title"));
    }
}
