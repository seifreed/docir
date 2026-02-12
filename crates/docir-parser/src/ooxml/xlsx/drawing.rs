use super::XlsxParser;
use crate::error::ParseError;
use crate::ooxml::relationships::{Relationships, TargetMode};
use crate::zip_handler::PackageReader;
use docir_core::ir::{IRNode, Shape, ShapeType, WorksheetDrawing};
use docir_core::security::{ExternalRefType, ExternalReference};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::events::Event;
use quick_xml::Reader;

impl XlsxParser {
    pub(super) fn parse_drawing(
        &mut self,
        xml: &str,
        drawing_path: &str,
        relationships: &Relationships,
        zip: &mut impl PackageReader,
    ) -> Result<NodeId, ParseError> {
        let mut drawing = WorksheetDrawing::new();
        drawing.span = Some(SourceSpan::new(drawing_path));

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut current_shape: Option<Shape> = None;
        let mut current_embed: Option<String> = None;
        let mut current_chart: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"xdr:pic" => {
                        current_shape = Some(Shape::new(ShapeType::Picture));
                    }
                    b"xdr:graphicFrame" => {
                        current_shape = Some(Shape::new(ShapeType::Chart));
                    }
                    b"xdr:cNvPr" => {
                        if let Some(shape) = current_shape.as_mut() {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"name" => {
                                        shape.name =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                    b"descr" => {
                                        shape.alt_text =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    b"a:blip" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r:embed" {
                                current_embed =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ if e.name().as_ref().ends_with(b"chart") => {
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            if key == b"r:id" || key.ends_with(b":id") {
                                current_chart =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => match e.name().as_ref() {
                    b"xdr:pic" => {
                        if let Some(mut shape) = current_shape.take() {
                            if let Some(rel_id) = current_embed.take() {
                                if let Some(rel) = relationships.get(&rel_id) {
                                    shape.relationship_id = Some(rel_id.clone());
                                    shape.media_target = Some(Relationships::resolve_target(
                                        drawing_path,
                                        &rel.target,
                                    ));
                                    if rel.target_mode == TargetMode::External {
                                        let ext_ref = ExternalReference::new(
                                            ExternalRefType::Image,
                                            &rel.target,
                                        );
                                        let ext_ref = ExternalReference {
                                            relationship_id: Some(rel_id),
                                            ..ext_ref
                                        };
                                        let ext_id = ext_ref.id;
                                        self.store.insert(IRNode::ExternalReference(ext_ref));
                                        self.security_info.external_refs.push(ext_id);
                                    }
                                }
                            }
                            let id = shape.id;
                            self.store.insert(IRNode::Shape(shape));
                            drawing.shapes.push(id);
                        }
                    }
                    b"xdr:graphicFrame" => {
                        if let Some(mut shape) = current_shape.take() {
                            if let Some(rel_id) = current_chart.take() {
                                if let Some(rel) = relationships.get(&rel_id) {
                                    shape.relationship_id = Some(rel_id.clone());
                                    shape.media_target = Some(Relationships::resolve_target(
                                        drawing_path,
                                        &rel.target,
                                    ));
                                    let chart_path =
                                        Relationships::resolve_target(drawing_path, &rel.target);
                                    if zip.contains(&chart_path) {
                                        let chart_xml = zip.read_file_string(&chart_path)?;
                                        if let Some(chart_id) =
                                            self.parse_chart(&chart_xml, &chart_path)
                                        {
                                            self.chart_nodes.push(chart_id);
                                        }
                                    }
                                }
                            }
                            let id = shape.id;
                            self.store.insert(IRNode::Shape(shape));
                            drawing.shapes.push(id);
                        }
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: drawing_path.to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        let id = drawing.id;
        self.store.insert(IRNode::WorksheetDrawing(drawing));
        Ok(id)
    }
}
