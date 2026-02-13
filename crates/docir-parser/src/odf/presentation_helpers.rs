//! ODF presentation parsing helpers extracted from the main module.

use super::*;

pub(super) struct FrameShapeState {
    pub(super) shape_type: ShapeType,
    pub(super) media_target: Option<String>,
    pub(super) chart_id: Option<NodeId>,
    pub(super) has_shape: bool,
}

impl FrameShapeState {
    pub(super) fn new() -> Self {
        Self {
            shape_type: ShapeType::Picture,
            media_target: None,
            chart_id: None,
            has_shape: false,
        }
    }
}

pub(super) fn parse_draw_page(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    slide_no: u32,
    store: &mut IrStore,
) -> Result<Slide, ParseError> {
    let mut slide = Slide::new(slide_no);
    slide.name = attr_value(start, b"draw:name");
    slide.master_id = attr_value(start, b"draw:master-page-name")
        .or_else(|| attr_value(start, b"presentation:master-page-name"));
    slide.layout_id = attr_value(start, b"presentation:page-layout-name")
        .or_else(|| attr_value(start, b"draw:style-name"));
    slide.transition = parse_odp_transition(start);
    let mut buf = Vec::new();
    let mut notes_text: Option<String> = None;

    loop {
        match read_event(reader, &mut buf, "content.xml")? {
            Event::Start(e) => match e.name().as_ref() {
                b"draw:frame" => {
                    if let Some(shape_id) = parse_draw_frame_presentation(reader, &e, store)? {
                        slide.shapes.push(shape_id);
                    }
                }
                b"draw:custom-shape" => {
                    if let Some(shape_id) = parse_custom_shape_presentation(reader, &e, store)? {
                        slide.shapes.push(shape_id);
                    }
                }
                b"presentation:notes" => {
                    notes_text = parse_notes(reader)?;
                }
                name if name.starts_with(b"anim:") => {
                    if let Some(anim) = parse_odf_animation(&e) {
                        slide.animations.push(anim);
                    }
                }
                _ => {}
            },
            Event::Empty(e) => match e.name().as_ref() {
                b"draw:frame" => {
                    if let Some(shape_id) = parse_draw_frame_presentation(reader, &e, store)? {
                        slide.shapes.push(shape_id);
                    }
                }
                b"draw:custom-shape" => {
                    if let Some(shape_id) = parse_custom_shape_presentation(reader, &e, store)? {
                        slide.shapes.push(shape_id);
                    }
                }
                name if name.starts_with(b"anim:") => {
                    if let Some(anim) = parse_odf_animation(&e) {
                        slide.animations.push(anim);
                    }
                }
                _ => {}
            },
            Event::End(e) => {
                if e.name().as_ref() == b"draw:page" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    slide.notes = notes_text;
    Ok(slide)
}

pub(super) fn parse_draw_frame_presentation(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let transform = parse_frame_transform(start);
    let mut frame = FrameShapeState::new();
    let mut text: Option<ShapeText> = None;
    let mut name = attr_value(start, b"draw:name");
    let mut buf = Vec::new();

    loop {
        match read_event(reader, &mut buf, "content.xml")? {
            Event::Start(e) => match e.name().as_ref() {
                b"draw:text-box" => {
                    let paragraphs = parse_shape_text(reader, b"draw:text-box")?;
                    if !paragraphs.is_empty() {
                        text = Some(ShapeText { paragraphs });
                        frame.shape_type = ShapeType::TextBox;
                        frame.has_shape = true;
                    }
                }
                _ => parse_frame_shape_start(reader, &e, store, &mut frame)?,
            },
            Event::Empty(e) => parse_frame_shape_empty(&e, store, &mut frame),
            Event::End(e) => {
                if e.name().as_ref() == b"draw:frame" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if frame.has_shape {
        let mut shape = Shape::new(frame.shape_type);
        shape.name = name.take();
        shape.media_target = frame.media_target;
        shape.text = text;
        shape.chart_id = frame.chart_id;
        shape.transform = transform;
        let shape_id = shape.id;
        store.insert(IRNode::Shape(shape));
        Ok(Some(shape_id))
    } else {
        Ok(None)
    }
}

pub(super) fn parse_custom_shape_presentation(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let mut name = attr_value(start, b"draw:name");
    let paragraphs = parse_shape_text(reader, b"draw:custom-shape")?;
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
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:p" => {
                    let text = parse_text_element(reader, b"text:p")?;
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
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_tag {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(paragraphs)
}

pub(super) fn parse_odp_transition(start: &BytesStart<'_>) -> Option<SlideTransition> {
    let transition_type = attr_value(start, b"presentation:transition-type")
        .or_else(|| attr_value(start, b"draw:transition-type"));
    let speed = attr_value(start, b"presentation:transition-speed");
    let duration_ms =
        attr_value(start, b"presentation:transition-duration").and_then(|v| v.parse::<u32>().ok());
    let advance_after_ms =
        attr_value(start, b"presentation:duration").and_then(|v| v.parse::<u32>().ok());
    let advance_on_click = attr_value(start, b"presentation:animation").map(|v| v == "click");
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

pub(super) fn parse_frame_shape_start(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
    frame: &mut FrameShapeState,
) -> Result<(), ParseError> {
    match start.name().as_ref() {
        b"draw:image" => apply_picture_shape(start, frame),
        b"chart:chart" => {
            frame.shape_type = ShapeType::Chart;
            frame.has_shape = true;
            let chart = parse_odf_chart(reader, start)?;
            let id = chart.id;
            store.insert(IRNode::ChartData(chart));
            frame.chart_id = Some(id);
        }
        b"draw:plugin" => apply_plugin_shape(start, frame),
        b"draw:object" | b"draw:object-ole" => apply_ole_shape(start, frame),
        _ => {}
    }
    Ok(())
}

pub(super) fn parse_frame_shape_empty(
    start: &BytesStart<'_>,
    store: &mut IrStore,
    frame: &mut FrameShapeState,
) {
    match start.name().as_ref() {
        b"draw:image" => apply_picture_shape(start, frame),
        b"chart:chart" => {
            frame.shape_type = ShapeType::Chart;
            frame.has_shape = true;
            let mut chart = ChartData::new();
            chart.chart_type = attr_value(start, b"chart:class");
            chart.span = Some(SourceSpan::new("content.xml"));
            let id = chart.id;
            store.insert(IRNode::ChartData(chart));
            frame.chart_id = Some(id);
        }
        b"draw:plugin" => apply_plugin_shape(start, frame),
        b"draw:object" | b"draw:object-ole" => apply_ole_shape(start, frame),
        _ => {}
    }
}

fn apply_picture_shape(start: &BytesStart<'_>, frame: &mut FrameShapeState) {
    if let Some(href) = attr_value(start, b"xlink:href") {
        frame.media_target = Some(href);
        frame.shape_type = ShapeType::Picture;
        frame.has_shape = true;
    }
}

fn apply_plugin_shape(start: &BytesStart<'_>, frame: &mut FrameShapeState) {
    if let Some(href) = attr_value(start, b"xlink:href") {
        frame.media_target = Some(href.clone());
        frame.shape_type = classify_media_shape(&href);
        frame.has_shape = true;
    }
}

fn apply_ole_shape(start: &BytesStart<'_>, frame: &mut FrameShapeState) {
    if let Some(href) = attr_value(start, b"xlink:href") {
        frame.media_target = Some(href);
    }
    frame.shape_type = ShapeType::OleObject;
    frame.has_shape = true;
}

pub(super) fn parse_odf_chart(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
) -> Result<ChartData, ParseError> {
    let mut chart = ChartData::new();
    chart.chart_type = attr_value(start, b"chart:class");
    chart.span = Some(SourceSpan::new("content.xml"));
    let mut buf = Vec::new();
    let mut in_title = false;
    let mut title_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"chart:title" => {
                    in_title = true;
                }
                b"text:p" if in_title => {
                    let text = parse_text_element(reader, b"text:p")?;
                    if !title_text.is_empty() && !text.is_empty() {
                        title_text.push(' ');
                    }
                    title_text.push_str(&text);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"chart:title" {
                    in_title = false;
                }
                if e.name().as_ref() == b"chart:chart" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
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
        target: attr_value(start, b"anim:targetElement"),
        duration_ms: attr_value(start, b"smil:dur").and_then(|v| parse_duration_ms(&v)),
        preset_id: attr_value(start, b"presentation:preset-id"),
        preset_class: attr_value(start, b"presentation:preset-class"),
        media_asset: None,
    };
    if anim.target.is_none() {
        anim.target = attr_value(start, b"smil:targetElement");
    }
    Some(anim)
}

fn parse_duration_ms(value: &str) -> Option<u32> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return None;
    }
    if let Some(stripped) = trimmed.strip_suffix("ms") {
        return stripped.parse::<u32>().ok();
    }
    if let Some(stripped) = trimmed.strip_suffix('s') {
        return stripped
            .parse::<f32>()
            .ok()
            .map(|v| (v * 1000.0).round() as u32);
    }
    if trimmed.starts_with("PT") && trimmed.ends_with('S') {
        let inner = trimmed.trim_start_matches("PT").trim_end_matches('S');
        return inner
            .parse::<f32>()
            .ok()
            .map(|v| (v * 1000.0).round() as u32);
    }
    None
}

pub(super) fn build_media_asset(path: &str, media: &str, size_bytes: u64) -> Option<MediaAsset> {
    let media_type = classify_media_type(path, media)?;
    let mut asset = MediaAsset::new(path.to_string(), media_type, size_bytes);
    asset.content_type = Some(media.to_string());
    asset.span = Some(SourceSpan::new("META-INF/manifest.xml"));
    Some(asset)
}

pub(super) fn classify_media_type(path: &str, media: &str) -> Option<MediaType> {
    let lower_media = media.to_ascii_lowercase();
    if lower_media.starts_with("image/") {
        return Some(MediaType::Image);
    }
    if lower_media.starts_with("audio/") {
        return Some(MediaType::Audio);
    }
    if lower_media.starts_with("video/") {
        return Some(MediaType::Video);
    }
    if lower_media.starts_with("application/") {
        let lower_path = path.to_ascii_lowercase();
        if lower_path.ends_with(".ogg") || lower_path.ends_with(".oga") {
            return Some(MediaType::Audio);
        }
        if lower_path.ends_with(".ogv") {
            return Some(MediaType::Video);
        }
    }
    None
}

pub(super) fn classify_media_shape(path: &str) -> ShapeType {
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".mp3")
        || lower.ends_with(".wav")
        || lower.ends_with(".ogg")
        || lower.ends_with(".oga")
    {
        return ShapeType::Audio;
    }
    if lower.ends_with(".mp4")
        || lower.ends_with(".mov")
        || lower.ends_with(".avi")
        || lower.ends_with(".mpeg")
        || lower.ends_with(".mpg")
        || lower.ends_with(".ogv")
    {
        return ShapeType::Video;
    }
    ShapeType::Unknown
}
