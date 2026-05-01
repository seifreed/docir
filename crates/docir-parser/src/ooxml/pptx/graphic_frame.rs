use super::graphic_frame_support::{
    apply_graphic_data_shape_type, apply_non_visual_shape_props, capture_chart_rel, capture_ole_rel,
};
use super::{
    parse_transform, BytesStart, ExternalRefType, ExternalReference, IRNode, NodeId, ParseError,
    PptxParser, Reader, Relationships, Shape, ShapeType, TargetMode,
};
use crate::zip_handler::PackageReader;

pub(super) struct GraphicFrameState {
    pub(super) chart_rel: Option<String>,
    pub(super) ole_rel: Option<String>,
    pub(super) table_id: Option<NodeId>,
}

impl GraphicFrameState {
    pub(super) fn new() -> Self {
        Self {
            chart_rel: None,
            ole_rel: None,
            table_id: None,
        }
    }
}

impl PptxParser {
    pub(super) fn handle_graphic_frame_start(
        &mut self,
        e: &BytesStart<'_>,
        reader: &mut Reader<&[u8]>,
        slide_path: &str,
        relationships: &Relationships,
        shape: &mut Shape,
        state: &mut GraphicFrameState,
    ) -> Result<(), ParseError> {
        match e.name().as_ref() {
            b"p:cNvPr" => apply_non_visual_shape_props(shape, e),
            b"a:hlinkClick" => {
                self.attach_hyperlink(shape, e, relationships, slide_path);
            }
            b"p:xfrm" => {
                parse_transform(reader, &mut shape.transform, slide_path)?;
            }
            _ if e.name().as_ref().ends_with(b"graphicData") => {
                apply_graphic_data_shape_type(shape, e);
            }
            b"a:tbl" => {
                let table = self.parse_pptx_table(reader, slide_path)?;
                let id = table.id;
                self.store.insert(IRNode::Table(table));
                state.table_id = Some(id);
                shape.shape_type = ShapeType::Table;
            }
            _ if e.name().as_ref().ends_with(b"chart") => {
                capture_chart_rel(state, e);
                shape.shape_type = ShapeType::Chart;
            }
            _ if e.name().as_ref().ends_with(b"oleObj")
                || e.name().as_ref().ends_with(b"oleObject") =>
            {
                capture_ole_rel(state, e);
                shape.shape_type = ShapeType::OleObject;
            }
            _ => {}
        }
        Ok(())
    }

    pub(super) fn handle_graphic_frame_empty(
        &mut self,
        e: &BytesStart<'_>,
        reader: &mut Reader<&[u8]>,
        slide_path: &str,
        relationships: &Relationships,
        shape: &mut Shape,
        state: &mut GraphicFrameState,
    ) -> Result<(), ParseError> {
        match e.name().as_ref() {
            b"p:cNvPr" => apply_non_visual_shape_props(shape, e),
            b"a:hlinkClick" => {
                self.attach_hyperlink(shape, e, relationships, slide_path);
            }
            b"p:xfrm" => {}
            _ if e.name().as_ref().ends_with(b"graphicData") => {}
            b"a:tbl" => {
                let table = self.parse_pptx_table(reader, slide_path)?;
                let id = table.id;
                self.store.insert(IRNode::Table(table));
                state.table_id = Some(id);
                shape.shape_type = ShapeType::Table;
            }
            _ if e.name().as_ref().ends_with(b"chart") => {
                capture_chart_rel(state, e);
                shape.shape_type = ShapeType::Chart;
            }
            _ if e.name().as_ref().ends_with(b"oleObj")
                || e.name().as_ref().ends_with(b"oleObject") =>
            {
                capture_ole_rel(state, e);
                shape.shape_type = ShapeType::OleObject;
            }
            _ => {}
        }
        Ok(())
    }

    pub(super) fn apply_graphic_frame_relationships(
        &mut self,
        shape: &mut Shape,
        slide_path: &str,
        relationships: &Relationships,
        zip: &mut impl PackageReader,
        state: &GraphicFrameState,
    ) -> Result<(), ParseError> {
        if let Some(rel_id) = state.ole_rel.as_ref() {
            if let Some(rel) = relationships.get(rel_id) {
                shape.shape_type = ShapeType::OleObject;
                shape.relationship_id = Some(rel_id.clone());
                let resolved = if rel.target_mode == TargetMode::External {
                    rel.target.clone()
                } else {
                    Relationships::resolve_target(slide_path, &rel.target)
                };
                shape.media_target = Some(resolved);
                if rel.target_mode == TargetMode::External {
                    let ext_ref = ExternalReference::new(ExternalRefType::Other, &rel.target);
                    let ext_ref = ExternalReference {
                        relationship_id: Some(rel_id.clone()),
                        ..ext_ref
                    };
                    let ext_id = ext_ref.id;
                    self.store.insert(IRNode::ExternalReference(ext_ref));
                    self.security_info.external_refs.push(ext_id);
                }
            }
            return Ok(());
        }

        if let Some(rel_id) = state.chart_rel.as_ref() {
            if let Some(rel) = relationships.get(rel_id) {
                if shape.shape_type == ShapeType::Custom && rel.rel_type.contains("chart") {
                    shape.shape_type = ShapeType::Chart;
                }
                shape.relationship_id = Some(rel_id.clone());
                let resolved = if rel.target_mode == TargetMode::External {
                    rel.target.clone()
                } else {
                    Relationships::resolve_target(slide_path, &rel.target)
                };
                shape.media_target = Some(resolved);
                if rel.target_mode == TargetMode::External {
                    let ext_ref = ExternalReference::new(ExternalRefType::Other, &rel.target);
                    let ext_ref = ExternalReference {
                        relationship_id: Some(rel_id.clone()),
                        ..ext_ref
                    };
                    let ext_id = ext_ref.id;
                    self.store.insert(IRNode::ExternalReference(ext_ref));
                    self.security_info.external_refs.push(ext_id);
                } else {
                    let chart_path = Relationships::resolve_target(slide_path, &rel.target);
                    if zip.contains(&chart_path) {
                        let chart_xml = zip.read_file_string(&chart_path)?;
                        if let Some(chart_id) = crate::ooxml::shared::parse_chart_data(
                            &chart_xml,
                            &chart_path,
                            &mut self.store,
                        ) {
                            self.chart_nodes.push(chart_id);
                        }
                    }
                }
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::relationships::{Relationship, TargetMode};
    use std::collections::HashMap;

    struct DummyZip;

    impl PackageReader for DummyZip {
        fn contains(&self, _name: &str) -> bool {
            false
        }

        fn read_file(&mut self, name: &str) -> Result<Vec<u8>, ParseError> {
            Err(ParseError::MissingPart(name.to_string()))
        }

        fn read_file_string(&mut self, name: &str) -> Result<String, ParseError> {
            Err(ParseError::MissingPart(name.to_string()))
        }

        fn file_size(&mut self, name: &str) -> Result<u64, ParseError> {
            Err(ParseError::MissingPart(name.to_string()))
        }

        fn file_names(&self) -> Vec<String> {
            Vec::new()
        }

        fn list_prefix(&self, _prefix: &str) -> Vec<String> {
            Vec::new()
        }

        fn list_suffix(&self, _suffix: &str) -> Vec<String> {
            Vec::new()
        }
    }

    fn relationships_with(rel: Relationship) -> Relationships {
        let mut by_id = HashMap::new();
        by_id.insert(rel.id.clone(), rel.clone());
        let mut by_type = HashMap::new();
        by_type.insert(rel.rel_type.clone(), vec![rel.id]);
        Relationships { by_id, by_type }
    }

    #[test]
    fn graphic_frame_state_new_initializes_empty_fields() {
        let state = GraphicFrameState::new();
        assert!(state.chart_rel.is_none());
        assert!(state.ole_rel.is_none());
        assert!(state.table_id.is_none());
    }

    #[test]
    fn apply_graphic_frame_relationships_maps_external_ole_reference() {
        let rel = Relationship {
            id: "rIdOle".to_string(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
                    .to_string(),
            target: "https://example.com/object.bin".to_string(),
            target_mode: TargetMode::External,
        };
        let relationships = relationships_with(rel);
        let mut parser = PptxParser::new();
        let mut shape = Shape::new(ShapeType::Custom);
        let state = GraphicFrameState {
            chart_rel: None,
            ole_rel: Some("rIdOle".to_string()),
            table_id: None,
        };
        let mut zip = DummyZip;

        let result = parser.apply_graphic_frame_relationships(
            &mut shape,
            "ppt/slides/slide1.xml",
            &relationships,
            &mut zip,
            &state,
        );
        assert!(result.is_ok());
        assert_eq!(shape.shape_type, ShapeType::OleObject);
        assert_eq!(shape.relationship_id.as_deref(), Some("rIdOle"));
        assert_eq!(
            shape.media_target.as_deref(),
            Some("https://example.com/object.bin")
        );
        assert_eq!(parser.security_info.external_refs.len(), 1);
    }

    #[test]
    fn apply_graphic_frame_relationships_maps_external_chart_reference() {
        let rel = Relationship {
            id: "rIdChart".to_string(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
                .to_string(),
            target: "https://example.com/chart.xml".to_string(),
            target_mode: TargetMode::External,
        };
        let relationships = relationships_with(rel);
        let mut parser = PptxParser::new();
        let mut shape = Shape::new(ShapeType::Custom);
        let state = GraphicFrameState {
            chart_rel: Some("rIdChart".to_string()),
            ole_rel: None,
            table_id: None,
        };
        let mut zip = DummyZip;

        let result = parser.apply_graphic_frame_relationships(
            &mut shape,
            "ppt/slides/slide1.xml",
            &relationships,
            &mut zip,
            &state,
        );
        assert!(result.is_ok());
        assert_eq!(shape.shape_type, ShapeType::Chart);
        assert_eq!(shape.relationship_id.as_deref(), Some("rIdChart"));
        assert_eq!(
            shape.media_target.as_deref(),
            Some("https://example.com/chart.xml")
        );
        assert_eq!(parser.security_info.external_refs.len(), 1);
    }

    #[test]
    fn apply_graphic_frame_relationships_resolves_internal_chart_target() {
        let rel = Relationship {
            id: "rIdChart".to_string(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
                .to_string(),
            target: "../charts/chart1.xml".to_string(),
            target_mode: TargetMode::Internal,
        };
        let relationships = relationships_with(rel);
        let mut parser = PptxParser::new();
        let mut shape = Shape::new(ShapeType::Custom);
        let state = GraphicFrameState {
            chart_rel: Some("rIdChart".to_string()),
            ole_rel: None,
            table_id: None,
        };
        let mut zip = DummyZip;

        let result = parser.apply_graphic_frame_relationships(
            &mut shape,
            "ppt/slides/slide1.xml",
            &relationships,
            &mut zip,
            &state,
        );
        assert!(result.is_ok());
        assert_eq!(shape.shape_type, ShapeType::Chart);
        assert_eq!(shape.media_target.as_deref(), Some("ppt/charts/chart1.xml"));
        assert!(parser.security_info.external_refs.is_empty());
        assert!(parser.chart_nodes.is_empty());
    }

    #[test]
    fn apply_graphic_frame_relationships_with_missing_ole_rel_id_is_noop() {
        let relationships = Relationships::default();
        let mut parser = PptxParser::new();
        let mut shape = Shape::new(ShapeType::Custom);
        let state = GraphicFrameState {
            chart_rel: None,
            ole_rel: Some("rIdMissingOle".to_string()),
            table_id: None,
        };
        let mut zip = DummyZip;

        let result = parser.apply_graphic_frame_relationships(
            &mut shape,
            "ppt/slides/slide1.xml",
            &relationships,
            &mut zip,
            &state,
        );
        assert!(result.is_ok());
        assert_eq!(shape.shape_type, ShapeType::Custom);
        assert!(shape.relationship_id.is_none());
        assert!(shape.media_target.is_none());
        assert!(parser.security_info.external_refs.is_empty());
    }

    #[test]
    fn apply_graphic_frame_relationships_with_missing_chart_rel_id_is_noop() {
        let relationships = Relationships::default();
        let mut parser = PptxParser::new();
        let mut shape = Shape::new(ShapeType::Custom);
        let state = GraphicFrameState {
            chart_rel: Some("rIdMissingChart".to_string()),
            ole_rel: None,
            table_id: None,
        };
        let mut zip = DummyZip;

        let result = parser.apply_graphic_frame_relationships(
            &mut shape,
            "ppt/slides/slide1.xml",
            &relationships,
            &mut zip,
            &state,
        );
        assert!(result.is_ok());
        assert_eq!(shape.shape_type, ShapeType::Custom);
        assert!(shape.relationship_id.is_none());
        assert!(shape.media_target.is_none());
        assert!(parser.security_info.external_refs.is_empty());
        assert!(parser.chart_nodes.is_empty());
    }
}
