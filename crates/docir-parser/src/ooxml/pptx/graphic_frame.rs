use super::*;
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
            b"p:cNvPr" => {
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"name" => {
                            shape.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                        b"descr" => {
                            shape.alt_text = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                        _ => {}
                    }
                }
            }
            b"a:hlinkClick" => {
                self.attach_hyperlink(shape, e, relationships, slide_path);
            }
            b"p:xfrm" => {
                parse_transform(reader, &mut shape.transform, slide_path)?;
            }
            _ if e.name().as_ref().ends_with(b"graphicData") => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"uri" {
                        let uri = String::from_utf8_lossy(&attr.value);
                        if uri.contains("chart") {
                            shape.shape_type = ShapeType::Chart;
                        } else if uri.contains("table") {
                            shape.shape_type = ShapeType::Table;
                        } else if uri.contains("ole") || uri.contains("object") {
                            shape.shape_type = ShapeType::OleObject;
                        }
                    }
                }
            }
            b"a:tbl" => {
                let table = self.parse_pptx_table(reader, slide_path)?;
                let id = table.id;
                self.store.insert(IRNode::Table(table));
                state.table_id = Some(id);
                shape.shape_type = ShapeType::Table;
            }
            _ if e.name().as_ref().ends_with(b"chart") => {
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    if state.chart_rel.is_none()
                        && (key == b"r:id" || key == b"id" || key.ends_with(b":id"))
                    {
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        if val.starts_with("rId") {
                            state.chart_rel = Some(val);
                        }
                    }
                }
                shape.shape_type = ShapeType::Chart;
            }
            _ if e.name().as_ref().ends_with(b"oleObj")
                || e.name().as_ref().ends_with(b"oleObject") =>
            {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"r:id" {
                        state.ole_rel = Some(String::from_utf8_lossy(&attr.value).to_string());
                    }
                }
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
            b"p:cNvPr" => {
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"name" => {
                            shape.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                        b"descr" => {
                            shape.alt_text = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                        _ => {}
                    }
                }
            }
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
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    if state.chart_rel.is_none()
                        && (key == b"r:id" || key == b"id" || key.ends_with(b":id"))
                    {
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        if val.starts_with("rId") {
                            state.chart_rel = Some(val);
                        }
                    }
                }
                shape.shape_type = ShapeType::Chart;
            }
            _ if e.name().as_ref().ends_with(b"oleObj")
                || e.name().as_ref().ends_with(b"oleObject") =>
            {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"r:id" {
                        state.ole_rel = Some(String::from_utf8_lossy(&attr.value).to_string());
                    }
                }
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
