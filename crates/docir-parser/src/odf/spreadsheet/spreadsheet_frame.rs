use super::super::super::presentation_helpers::{
    parse_frame_shape_empty, parse_frame_shape_start, FrameShapeState,
};
use super::super::super::{
    parse_frame_transform, IRNode, IrStore, NodeId, OdfReader, ParseError, Shape,
};
use crate::xml_utils::attr_value;
use crate::xml_utils::{is_end_event, scan_xml_events_until_end, XmlScanControl};
use quick_xml::events::{BytesStart, Event};

pub(crate) fn parse_draw_frame_spreadsheet(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let transform = parse_frame_transform(start);
    let mut frame = FrameShapeState::new();
    let mut buf = Vec::new();
    let mut name = attr_value(start, b"draw:name");

    scan_xml_events_until_end(
        reader,
        &mut buf,
        "content.xml",
        |event| is_end_event(event, b"draw:frame"),
        |reader, event| {
            match event {
                Event::Start(start) => {
                    parse_frame_shape_start(reader, start, store, &mut frame)?;
                }
                Event::Empty(start) => {
                    parse_frame_shape_empty(start, store, &mut frame);
                }
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        },
    )?;

    if frame.has_shape {
        let mut shape = Shape::new(frame.shape_type);
        shape.name = name.take();
        shape.media_target = frame.media_target;
        shape.chart_id = frame.chart_id;
        shape.transform = transform;
        let shape_id = shape.id;
        store.insert(IRNode::Shape(shape));
        Ok(Some(shape_id))
    } else {
        Ok(None)
    }
}
