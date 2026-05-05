use super::{
    parse_odf_chart, ChartData, IRNode, IrStore, NodeId, OdfReader, ParseError, ShapeType,
    SourceSpan,
};
use crate::xml_utils::{attr_value_by_suffix, local_name};
use quick_xml::events::BytesStart;

pub(crate) struct FrameShapeState {
    pub(crate) shape_type: ShapeType,
    pub(crate) media_target: Option<String>,
    pub(crate) chart_id: Option<NodeId>,
    pub(crate) has_shape: bool,
}

impl FrameShapeState {
    pub(crate) fn new() -> Self {
        Self {
            shape_type: ShapeType::Picture,
            media_target: None,
            chart_id: None,
            has_shape: false,
        }
    }
}

pub(crate) fn parse_frame_shape_start(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
    frame: &mut FrameShapeState,
) -> Result<(), ParseError> {
    match local_name(start.name().as_ref()) {
        b"image" => apply_picture_shape(start, frame),
        b"chart" => {
            frame.shape_type = ShapeType::Chart;
            frame.has_shape = true;
            let chart = parse_odf_chart(reader, start)?;
            let id = chart.id;
            store.insert(IRNode::ChartData(chart));
            frame.chart_id = Some(id);
        }
        b"plugin" => apply_plugin_shape(start, frame),
        b"object" | b"object-ole" => apply_ole_shape(start, frame),
        _ => {}
    }
    Ok(())
}

pub(crate) fn parse_frame_shape_empty(
    start: &BytesStart<'_>,
    store: &mut IrStore,
    frame: &mut FrameShapeState,
) {
    match local_name(start.name().as_ref()) {
        b"image" => apply_picture_shape(start, frame),
        b"chart" => {
            frame.shape_type = ShapeType::Chart;
            frame.has_shape = true;
            let mut chart = ChartData::new();
            chart.chart_type = attr_value_by_suffix(start, &[b":class"]);
            chart.span = Some(SourceSpan::new("content.xml"));
            let id = chart.id;
            store.insert(IRNode::ChartData(chart));
            frame.chart_id = Some(id);
        }
        b"plugin" => apply_plugin_shape(start, frame),
        b"object" | b"object-ole" => apply_ole_shape(start, frame),
        _ => {}
    }
}

fn apply_picture_shape(start: &BytesStart<'_>, frame: &mut FrameShapeState) {
    if let Some(href) = attr_value_by_suffix(start, &[b":href"]) {
        frame.media_target = Some(href);
        frame.shape_type = ShapeType::Picture;
        frame.has_shape = true;
    }
}

fn apply_plugin_shape(start: &BytesStart<'_>, frame: &mut FrameShapeState) {
    if let Some(href) = attr_value_by_suffix(start, &[b":href"]) {
        frame.media_target = Some(href.clone());
        frame.shape_type = classify_media_shape(&href);
        frame.has_shape = true;
    }
}

fn apply_ole_shape(start: &BytesStart<'_>, frame: &mut FrameShapeState) {
    if let Some(href) = attr_value_by_suffix(start, &[b":href"]) {
        frame.media_target = Some(href);
    }
    frame.shape_type = ShapeType::OleObject;
    frame.has_shape = true;
}

pub(crate) fn classify_media_shape(path: &str) -> ShapeType {
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
