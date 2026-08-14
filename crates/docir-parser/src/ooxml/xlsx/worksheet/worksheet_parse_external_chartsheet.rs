use crate::ooxml::relationships::Relationships;
use crate::ooxml::xlsx::{IRNode, ParseError, Shape, ShapeType, WorksheetDrawing, XlsxParser};
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::{
    XmlScanControl, local_name, reader_from_str, scan_xml_events, track_xml_document_event,
    visit_attributes,
};
use crate::zip_handler::PackageReader;
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::events::Event;

pub(super) fn parse_chartsheet_impl(
    parser: &mut XlsxParser,
    zip: &mut impl PackageReader,
    xml: &str,
    sheet_path: &str,
    relationships: &Relationships,
) -> Result<Option<NodeId>, ParseError> {
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let mut chart_rel: Option<String> = None;
    let mut depth = 0usize;
    let mut root_closed = false;

    scan_xml_events(&mut reader, &mut buf, sheet_path, |event| {
        track_xml_document_event(&event, &mut depth, &mut root_closed, sheet_path)?;
        match event {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"chart" => {
                visit_attributes(&e, sheet_path, |attr| {
                    if local_name(attr.key.as_ref()) == b"id" {
                        chart_rel = Some(lossy_attr_value(attr).to_string());
                    }
                })?;
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    let Some(rel_id) = chart_rel else {
        return Ok(None);
    };
    let Some(rel) = relationships.get(&rel_id) else {
        return Ok(None);
    };
    let chart_path = Relationships::resolve_target(sheet_path, &rel.target);
    if !zip.contains(&chart_path) {
        return Ok(None);
    }

    let chart_xml = zip.read_file_string(&chart_path)?;
    let chart_id = parser.parse_chart(&chart_xml, &chart_path)?;

    let mut drawing = WorksheetDrawing::new();
    drawing.span = Some(SourceSpan::new(sheet_path));
    let mut shape = Shape::new(ShapeType::Chart);
    shape.media_target = Some(chart_path);
    shape.relationship_id = Some(rel_id);
    shape.span = Some(SourceSpan::new(sheet_path));
    let shape_id = shape.id;
    parser.store.insert(IRNode::Shape(shape));
    drawing.shapes.push(shape_id);

    parser.chart_nodes.push(chart_id);

    let drawing_id = drawing.id;
    parser.store.insert(IRNode::WorksheetDrawing(drawing));
    Ok(Some(drawing_id))
}
