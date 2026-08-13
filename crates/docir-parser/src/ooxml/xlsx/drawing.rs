use super::XlsxParser;
use crate::error::ParseError;
use crate::ooxml::relationships::{Relationships, TargetMode};
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::{local_name, visit_attributes, xml_error};
use crate::zip_handler::PackageReader;
use docir_core::ir::{IRNode, Shape, ShapeType, WorksheetDrawing};
use docir_core::security::{ExternalRefType, ExternalReference};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

fn drawing_relationship_target(
    drawing_path: &str,
    target_mode: TargetMode,
    target: &str,
) -> String {
    if target_mode == TargetMode::External {
        target.to_string()
    } else {
        Relationships::resolve_target(drawing_path, target)
    }
}

impl XlsxParser {
    pub(super) fn parse_drawing(
        &mut self,
        xml: &str,
        drawing_path: &str,
        relationships: &Relationships,
        zip: &mut impl PackageReader,
    ) -> Result<NodeId, ParseError> {
        let mut state = XlsxDrawingState::new(drawing_path);

        let mut reader = Reader::from_str(xml);
        let config = reader.config_mut();
        config.trim_text(true);
        config.check_end_names = true;
        let mut buf = Vec::new();
        let mut depth = 0usize;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    depth += 1;
                    handle_xlsx_drawing_start(&e, &mut state, drawing_path)?;
                }
                Ok(Event::Empty(e)) => {
                    handle_xlsx_drawing_empty(&e, &mut state, drawing_path)?;
                }
                Ok(Event::End(e)) => {
                    depth = depth.saturating_sub(1);
                    match local_name(e.name().as_ref()) {
                        b"pic" => {
                            self.finish_picture_shape(&mut state, drawing_path, relationships);
                        }
                        b"graphicFrame" => {
                            self.finish_chart_shape(&mut state, drawing_path, relationships, zip)?;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) if depth == 0 => break,
                Ok(Event::Eof) => {
                    return Err(xml_error(drawing_path, "unexpected EOF in drawing XML"));
                }
                Err(e) => {
                    return Err(xml_error(drawing_path, e));
                }
                _ => {}
            }
            buf.clear();
        }

        let id = state.drawing.id;
        self.store.insert(IRNode::WorksheetDrawing(state.drawing));
        Ok(id)
    }

    fn finish_picture_shape(
        &mut self,
        state: &mut XlsxDrawingState,
        drawing_path: &str,
        relationships: &Relationships,
    ) {
        if let Some(mut shape) = state.current_shape.take() {
            if let Some(rel_id) = state.current_embed.take()
                && let Some(rel) = relationships.get(&rel_id)
            {
                shape.relationship_id = Some(rel_id.clone());
                shape.media_target = Some(drawing_relationship_target(
                    drawing_path,
                    rel.target_mode,
                    &rel.target,
                ));
                if rel.target_mode == TargetMode::External {
                    let ext_ref = ExternalReference::new(ExternalRefType::Image, &rel.target);
                    let ext_ref = ExternalReference {
                        relationship_id: Some(rel_id),
                        ..ext_ref
                    };
                    let ext_id = ext_ref.id;
                    self.store.insert(IRNode::ExternalReference(ext_ref));
                    self.security_info.external_refs.push(ext_id);
                }
            }
            state.insert_shape(shape, &mut self.store);
        }
    }

    fn finish_chart_shape(
        &mut self,
        state: &mut XlsxDrawingState,
        drawing_path: &str,
        relationships: &Relationships,
        zip: &mut impl PackageReader,
    ) -> Result<(), ParseError> {
        if let Some(mut shape) = state.current_shape.take() {
            if let Some(rel_id) = state.current_chart.take()
                && let Some(rel) = relationships.get(&rel_id)
            {
                shape.relationship_id = Some(rel_id.clone());
                let chart_path =
                    drawing_relationship_target(drawing_path, rel.target_mode, &rel.target);
                shape.media_target = Some(chart_path.clone());
                if rel.target_mode != TargetMode::External && zip.contains(&chart_path) {
                    let chart_xml = zip.read_file_string(&chart_path)?;
                    let chart_id = self.parse_chart(&chart_xml, &chart_path)?;
                    self.chart_nodes.push(chart_id);
                }
            }
            state.insert_shape(shape, &mut self.store);
        }
        Ok(())
    }
}

struct XlsxDrawingState {
    drawing: WorksheetDrawing,
    current_shape: Option<Shape>,
    current_embed: Option<String>,
    current_chart: Option<String>,
}

impl XlsxDrawingState {
    fn new(drawing_path: &str) -> Self {
        let mut drawing = WorksheetDrawing::new();
        drawing.span = Some(SourceSpan::new(drawing_path));
        Self {
            drawing,
            current_shape: None,
            current_embed: None,
            current_chart: None,
        }
    }

    fn insert_shape(&mut self, shape: Shape, store: &mut docir_core::visitor::IrStore) {
        let id = shape.id;
        store.insert(IRNode::Shape(shape));
        self.drawing.shapes.push(id);
    }
}

fn handle_xlsx_drawing_start(
    e: &BytesStart<'_>,
    state: &mut XlsxDrawingState,
    drawing_path: &str,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"pic" => state.current_shape = Some(Shape::new(ShapeType::Picture)),
        b"graphicFrame" => state.current_shape = Some(Shape::new(ShapeType::Chart)),
        b"cNvPr" => apply_shape_properties(e, state.current_shape.as_mut(), drawing_path)?,
        b"blip" => state.current_embed = relationship_attr(e, b"embed", drawing_path)?,
        b"chart" => state.current_chart = relationship_attr(e, b"id", drawing_path)?,
        _ => {}
    }
    Ok(())
}

fn handle_xlsx_drawing_empty(
    e: &BytesStart<'_>,
    state: &mut XlsxDrawingState,
    drawing_path: &str,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"cNvPr" => apply_shape_properties(e, state.current_shape.as_mut(), drawing_path)?,
        b"blip" => state.current_embed = relationship_attr(e, b"embed", drawing_path)?,
        b"chart" => state.current_chart = relationship_attr(e, b"id", drawing_path)?,
        _ => {}
    }
    Ok(())
}

fn apply_shape_properties(
    e: &BytesStart<'_>,
    shape: Option<&mut Shape>,
    drawing_path: &str,
) -> Result<(), ParseError> {
    if let Some(shape) = shape {
        visit_attributes(e, drawing_path, |attr| match attr.key.as_ref() {
            b"name" => shape.name = Some(lossy_attr_value(attr).to_string()),
            b"descr" => shape.alt_text = Some(lossy_attr_value(attr).to_string()),
            _ => {}
        })?;
    }
    Ok(())
}

fn relationship_attr(
    e: &BytesStart<'_>,
    local: &[u8],
    drawing_path: &str,
) -> Result<Option<String>, ParseError> {
    let mut value = None;
    visit_attributes(e, drawing_path, |attr| {
        if value.is_none() && local_name(attr.key.as_ref()) == local {
            value = Some(lossy_attr_value(attr).to_string());
        }
    })?;
    Ok(value)
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
            <xdr:wsDr xmlns:xdr="xdr" xmlns:a="a" xmlns:rel="r">
              <xdr:pic>
                <xdr:nvPicPr>
                  <xdr:cNvPr name="Picture 1" descr="Alt"/>
                </xdr:nvPicPr>
                <xdr:blipFill>
                  <a:blip rel:embed="rImgExternal"></a:blip>
                </xdr:blipFill>
              </xdr:pic>
              <xdr:pic>
                <xdr:nvPicPr>
                  <xdr:cNvPr name="Picture 2"></xdr:cNvPr>
                </xdr:nvPicPr>
                <xdr:blipFill>
                  <a:blip rel:embed="rImgInternal"></a:blip>
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
        let first_shape = parser
            .store
            .get(drawing.shapes[0])
            .and_then(|node| match node {
                IRNode::Shape(shape) => Some(shape),
                _ => None,
            })
            .expect("first shape");
        assert_eq!(first_shape.name.as_deref(), Some("Picture 1"));
        assert_eq!(first_shape.alt_text.as_deref(), Some("Alt"));
        assert_eq!(
            first_shape.media_target.as_deref(),
            Some("https://cdn.example.test/image.png")
        );

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
    fn parse_drawing_accepts_alternate_xml_prefixes() {
        let mut parser = XlsxParser::new();
        let rels = Relationships::parse(relationships_xml()).expect("relationships");
        let mut zip = TestPackageReader::new(&[(
            "xl/charts/chart1.xml",
            br#"<chartSpace><chart></chart></chartSpace>"#,
        )]);
        let drawing_xml = r#"
            <d:wsDr xmlns:d="drawing" xmlns:b="blip" xmlns:c="chart" xmlns:rel="rels">
              <d:pic>
                <d:nvPicPr>
                  <d:cNvPr name="Picture 1" descr="Alt"></d:cNvPr>
                </d:nvPicPr>
                <d:blipFill>
                  <b:blip rel:embed="rImgExternal"></b:blip>
                </d:blipFill>
              </d:pic>
              <d:graphicFrame>
                <d:nvGraphicFramePr>
                  <d:cNvPr name="Chart 1" descr="Chart alt"></d:cNvPr>
                </d:nvGraphicFramePr>
                <d:graphic>
                  <d:graphicData>
                    <c:chart rel:id="rChart"></c:chart>
                  </d:graphicData>
                </d:graphic>
              </d:graphicFrame>
            </d:wsDr>
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
        let chart_shape = parser
            .store
            .get(drawing.shapes[1])
            .and_then(|node| match node {
                IRNode::Shape(shape) => Some(shape),
                _ => None,
            })
            .expect("chart shape");
        assert_eq!(
            chart_shape.media_target.as_deref(),
            Some("xl/charts/chart1.xml")
        );
    }

    #[test]
    fn parse_drawing_reports_truncated_xml() {
        let mut parser = XlsxParser::new();
        let rels = Relationships::parse(relationships_xml()).expect("relationships");
        let mut zip = TestPackageReader::new(&[]);

        let err = parser
            .parse_drawing(
                "<xdr:wsDr><xdr:pic>",
                "xl/drawings/drawing1.xml",
                &rels,
                &mut zip,
            )
            .expect_err("truncated drawing XML must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/drawings/drawing1.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_drawing_reports_mismatched_xml() {
        let mut parser = XlsxParser::new();
        let rels = Relationships::parse(relationships_xml()).expect("relationships");
        let mut zip = TestPackageReader::new(&[]);

        let err = parser
            .parse_drawing(
                "<xdr:wsDr><xdr:pic></xdr:picx></xdr:wsDr>",
                "xl/drawings/drawing1.xml",
                &rels,
                &mut zip,
            )
            .expect_err("mismatched drawing XML must fail");
        assert!(matches!(err, ParseError::Xml { file, .. } if file == "xl/drawings/drawing1.xml"));
    }

    #[test]
    fn parse_drawing_reports_malformed_attributes() {
        let mut parser = XlsxParser::new();
        let rels = Relationships::parse(relationships_xml()).expect("relationships");
        let mut zip = TestPackageReader::new(&[]);
        let drawing_xml = r#"
            <xdr:wsDr xmlns:xdr="xdr">
              <xdr:pic>
                <xdr:nvPicPr>
                  <xdr:cNvPr name="Picture 1" name="Duplicate"/>
                </xdr:nvPicPr>
              </xdr:pic>
            </xdr:wsDr>
        "#;

        let err = parser
            .parse_drawing(drawing_xml, "xl/drawings/drawing1.xml", &rels, &mut zip)
            .expect_err("duplicate drawing attributes must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/drawings/drawing1.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }
}
