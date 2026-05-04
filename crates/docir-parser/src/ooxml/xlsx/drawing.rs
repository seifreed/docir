use super::XlsxParser;
use crate::error::ParseError;
use crate::ooxml::relationships::{Relationships, TargetMode};
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::xml_error;
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
                                        shape.name = Some(lossy_attr_value(&attr).to_string());
                                    }
                                    b"descr" => {
                                        shape.alt_text = Some(lossy_attr_value(&attr).to_string());
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    b"a:blip" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r:embed" {
                                current_embed = Some(lossy_attr_value(&attr).to_string());
                            }
                        }
                    }
                    _ if e.name().as_ref().ends_with(b"chart") => {
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            if key == b"r:id" || key.ends_with(b":id") {
                                current_chart = Some(lossy_attr_value(&attr).to_string());
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
                    return Err(xml_error(drawing_path, e));
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::xlsx::XlsxParser;
    use std::collections::HashMap;

    struct TestPackageReader {
        files: HashMap<String, Vec<u8>>,
    }

    impl TestPackageReader {
        fn new(entries: &[(&str, &[u8])]) -> Self {
            let files = entries
                .iter()
                .map(|(path, bytes)| ((*path).to_string(), bytes.to_vec()))
                .collect();
            Self { files }
        }
    }

    impl PackageReader for TestPackageReader {
        fn contains(&self, name: &str) -> bool {
            self.files.contains_key(name)
        }

        fn read_file(&mut self, name: &str) -> Result<Vec<u8>, ParseError> {
            self.files
                .get(name)
                .cloned()
                .ok_or_else(|| ParseError::MissingPart(name.to_string()))
        }

        fn file_size(&mut self, name: &str) -> Result<u64, ParseError> {
            self.files
                .get(name)
                .map(|v| v.len() as u64)
                .ok_or_else(|| ParseError::MissingPart(name.to_string()))
        }

        fn file_names(&self) -> Vec<String> {
            self.files.keys().cloned().collect()
        }

        fn list_prefix(&self, prefix: &str) -> Vec<String> {
            self.files
                .keys()
                .filter(|name| name.starts_with(prefix))
                .cloned()
                .collect()
        }

        fn list_suffix(&self, suffix: &str) -> Vec<String> {
            self.files
                .keys()
                .filter(|name| name.ends_with(suffix))
                .cloned()
                .collect()
        }
    }

    fn relationships_xml() -> &'static str {
        r#"
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rImgInternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
              <Relationship Id="rImgExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://cdn.example.test/image.png" TargetMode="External"/>
              <Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
            </Relationships>
        "#
    }

    #[test]
    fn parse_drawing_collects_shapes_and_tracks_external_reference() {
        let mut parser = XlsxParser::new();
        let rels = Relationships::parse(relationships_xml()).expect("relationships");
        let mut zip = TestPackageReader::new(&[]);
        let drawing_xml = r#"
            <xdr:wsDr xmlns:xdr="xdr" xmlns:a="a" xmlns:r="r">
              <xdr:pic>
                <xdr:nvPicPr>
                  <xdr:cNvPr name="Picture 1" descr="Alt"/>
                </xdr:nvPicPr>
                <xdr:blipFill>
                  <a:blip r:embed="rImgExternal"></a:blip>
                </xdr:blipFill>
              </xdr:pic>
              <xdr:pic>
                <xdr:nvPicPr>
                  <xdr:cNvPr name="Picture 2"></xdr:cNvPr>
                </xdr:nvPicPr>
                <xdr:blipFill>
                  <a:blip r:embed="rImgInternal"></a:blip>
                </xdr:blipFill>
              </xdr:pic>
            </xdr:wsDr>
        "#;

        let id = parser
            .parse_drawing(drawing_xml, "xl/drawings/drawing1.xml", &rels, &mut zip)
            .expect("drawing parse");

        let drawing = parser
            .store
            .get(id)
            .and_then(|node| match node {
                IRNode::WorksheetDrawing(d) => Some(d),
                _ => None,
            })
            .expect("worksheet drawing node");
        assert_eq!(drawing.shapes.len(), 2);

        let external_refs = parser
            .store
            .values()
            .filter(|node| matches!(node, IRNode::ExternalReference(_)))
            .count();
        assert_eq!(external_refs, 1);
        assert_eq!(parser.security_info.external_refs.len(), 1);
    }

    #[test]
    fn parse_drawing_reads_chart_relation_and_tolerates_unavailable_chart_payload() {
        let mut parser = XlsxParser::new();
        let rels = Relationships::parse(relationships_xml()).expect("relationships");
        let mut zip = TestPackageReader::new(&[(
            "xl/charts/chart1.xml",
            br#"<c:chartSpace xmlns:c="c"><c:chart></c:chart></c:chartSpace>"#,
        )]);
        let drawing_xml = r#"
            <xdr:wsDr xmlns:xdr="xdr" xmlns:c="c" xmlns:r="r">
              <xdr:graphicFrame>
                <xdr:nvGraphicFramePr>
                  <xdr:cNvPr name="Chart 1" descr="Chart alt"></xdr:cNvPr>
                </xdr:nvGraphicFramePr>
                <a:graphic xmlns:a="a">
                  <a:graphicData>
                    <c:chart r:id="rChart"></c:chart>
                  </a:graphicData>
                </a:graphic>
              </xdr:graphicFrame>
            </xdr:wsDr>
        "#;

        let id = parser
            .parse_drawing(drawing_xml, "xl/drawings/drawing1.xml", &rels, &mut zip)
            .expect("drawing parse");

        let drawing = parser
            .store
            .get(id)
            .and_then(|node| match node {
                IRNode::WorksheetDrawing(d) => Some(d),
                _ => None,
            })
            .expect("worksheet drawing node");
        assert_eq!(drawing.shapes.len(), 1);
    }

    #[test]
    fn parse_drawing_tolerates_truncated_xml_and_returns_empty_drawing() {
        let mut parser = XlsxParser::new();
        let rels = Relationships::parse(relationships_xml()).expect("relationships");
        let mut zip = TestPackageReader::new(&[]);

        let id = parser
            .parse_drawing(
                "<xdr:wsDr><xdr:pic>",
                "xl/drawings/drawing1.xml",
                &rels,
                &mut zip,
            )
            .expect("parser is tolerant for truncated drawing xml");
        let drawing = parser
            .store
            .get(id)
            .and_then(|node| match node {
                IRNode::WorksheetDrawing(d) => Some(d),
                _ => None,
            })
            .expect("worksheet drawing node");
        assert!(drawing.shapes.is_empty());
    }
}
