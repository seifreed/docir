use super::{
    BytesStart, Event, ExternalRefType, ParseError, PptxParser, Reader, Relationship,
    Relationships, Shape, ShapeType, SourceSpan, TargetMode, parse_shape_properties,
};
use crate::xml_utils::{local_name, lossy_attr_value, xml_error};

impl PptxParser {
    pub(super) fn parse_shape_pic(
        &mut self,
        reader: &mut Reader<&[u8]>,
        _start: &BytesStart,
        slide_path: &str,
        relationships: &Relationships,
    ) -> Result<Shape, ParseError> {
        let mut shape = Shape::new(ShapeType::Picture);
        shape.span = Some(SourceSpan::new(slide_path));

        let mut buf = Vec::new();
        let mut embed_rel: Option<String> = None;
        let mut link_rel: Option<String> = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    self.apply_picture_common_event(
                        &e,
                        &mut shape,
                        relationships,
                        slide_path,
                        &mut embed_rel,
                        &mut link_rel,
                    )?;
                    if local_name(e.name().as_ref()) == b"spPr" {
                        parse_shape_properties(reader, &mut shape, slide_path)?;
                    }
                }
                Ok(Event::Empty(e)) => {
                    self.apply_picture_common_event(
                        &e,
                        &mut shape,
                        relationships,
                        slide_path,
                        &mut embed_rel,
                        &mut link_rel,
                    )?;
                }
                Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"pic" => {
                    break;
                }
                Ok(Event::Eof) => {
                    return Err(xml_error(slide_path, "unexpected EOF in picture XML"));
                }
                Err(e) => {
                    return Err(xml_error(slide_path, e));
                }
                _ => {}
            }
            buf.clear();
        }

        self.apply_picture_relationships(
            &mut shape,
            relationships,
            slide_path,
            embed_rel,
            link_rel,
        );

        Ok(shape)
    }

    fn apply_picture_common_event(
        &mut self,
        event: &BytesStart<'_>,
        shape: &mut Shape,
        relationships: &Relationships,
        slide_path: &str,
        embed_rel: &mut Option<String>,
        link_rel: &mut Option<String>,
    ) -> Result<(), ParseError> {
        match local_name(event.name().as_ref()) {
            b"cNvPr" => apply_picture_non_visual_properties(event, shape, slide_path)?,
            b"hlinkClick" => {
                self.attach_hyperlink(shape, event, relationships, slide_path)?;
            }
            b"blip" => capture_picture_relationship_ids(event, embed_rel, link_rel, slide_path)?,
            _ => {}
        }
        Ok(())
    }

    fn apply_picture_relationships(
        &mut self,
        shape: &mut Shape,
        relationships: &Relationships,
        slide_path: &str,
        embed_rel: Option<String>,
        link_rel: Option<String>,
    ) {
        let primary_rel = embed_rel.as_ref().or(link_rel.as_ref()).cloned();
        if let Some(rel_id) = primary_rel {
            self.apply_primary_picture_relationship(shape, relationships, slide_path, rel_id);
        }
        self.add_linked_external_picture_reference(
            relationships,
            slide_path,
            embed_rel.as_deref(),
            link_rel.as_deref(),
        );
    }

    fn apply_primary_picture_relationship(
        &mut self,
        shape: &mut Shape,
        relationships: &Relationships,
        slide_path: &str,
        rel_id: String,
    ) {
        let Some(rel) = relationships.get(&rel_id) else {
            return;
        };

        shape.relationship_id = Some(rel_id);
        shape.media_target = Some(resolved_picture_target(rel, slide_path));
        if rel.rel_type.contains("audio") {
            shape.shape_type = ShapeType::Audio;
        } else if rel.rel_type.contains("video") {
            shape.shape_type = ShapeType::Video;
        }
        if rel.target_mode == TargetMode::External {
            self.add_external_reference(rel, picture_external_ref_type(rel), slide_path);
        }
    }

    fn add_linked_external_picture_reference(
        &mut self,
        relationships: &Relationships,
        slide_path: &str,
        embed_rel: Option<&str>,
        link_rel: Option<&str>,
    ) {
        let (Some(embed_id), Some(link_id)) = (embed_rel, link_rel) else {
            return;
        };
        if embed_id == link_id {
            return;
        }
        let Some(rel) = relationships.get(link_id) else {
            return;
        };
        if rel.target_mode == TargetMode::External {
            self.add_external_reference(rel, picture_external_ref_type(rel), slide_path);
        }
    }
}

fn apply_picture_non_visual_properties(
    event: &BytesStart<'_>,
    shape: &mut Shape,
    slide_path: &str,
) -> Result<(), ParseError> {
    for attr in event.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
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
    Ok(())
}

fn capture_picture_relationship_ids(
    event: &BytesStart<'_>,
    embed_rel: &mut Option<String>,
    link_rel: &mut Option<String>,
    slide_path: &str,
) -> Result<(), ParseError> {
    for attr in event.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
        match local_name(attr.key.as_ref()) {
            b"embed" => {
                *embed_rel = Some(lossy_attr_value(&attr).to_string());
            }
            b"link" => {
                *link_rel = Some(lossy_attr_value(&attr).to_string());
            }
            _ => {}
        }
    }
    Ok(())
}

fn resolved_picture_target(rel: &Relationship, slide_path: &str) -> String {
    if rel.target_mode == TargetMode::External {
        rel.target.clone()
    } else {
        Relationships::resolve_target(slide_path, &rel.target)
    }
}

fn picture_external_ref_type(rel: &Relationship) -> ExternalRefType {
    if rel.rel_type.contains("audio") || rel.rel_type.contains("video") {
        ExternalRefType::Other
    } else {
        ExternalRefType::Image
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_picture_fixture(xml: &str) -> Result<Shape, ParseError> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"pic" => {
                    let mut parser = PptxParser::new();
                    return parser.parse_shape_pic(
                        &mut reader,
                        &e,
                        "ppt/slides/broken-picture.xml",
                        &Relationships::default(),
                    );
                }
                Ok(Event::Eof) => {
                    return Err(xml_error("ppt/slides/broken-picture.xml", "missing pic"));
                }
                Err(err) => return Err(xml_error("ppt/slides/broken-picture.xml", err)),
                _ => {}
            }
            buf.clear();
        }
    }

    fn assert_picture_xml_error(result: Result<Shape, ParseError>) {
        match result.expect_err("malformed picture XML must fail") {
            ParseError::Xml { file, .. } => assert_eq!(file, "ppt/slides/broken-picture.xml"),
            other => panic!("expected XML error, got {other:?}"),
        }
    }

    #[test]
    fn parse_picture_reports_malformed_non_visual_attrs() {
        let xml = r#"
            <p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:nvPicPr><p:cNvPr name="Picture 1" name="Duplicate"/></p:nvPicPr>
            </p:pic>
        "#;

        assert_picture_xml_error(parse_picture_fixture(xml));
    }

    #[test]
    fn parse_picture_reports_malformed_blip_attrs() {
        let xml = r#"
            <p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:blipFill><a:blip r:embed="rId1" r:embed="rId2"/></p:blipFill>
            </p:pic>
        "#;

        assert_picture_xml_error(parse_picture_fixture(xml));
    }
}
