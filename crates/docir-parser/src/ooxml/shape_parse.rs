use super::{
    PackageReader, ParseError, PptxParser, Reader, Relationships, Shape, ShapeType, SourceSpan,
    TargetMode, classify_relationship, parse_shape_properties, parse_text_body, parse_transform,
    read_event,
};
use crate::xml_utils::{local_name, try_attr_value, try_attr_value_by_suffix, xml_error};
use docir_core::ir::IRNode;
use docir_core::types::NodeId;
use quick_xml::events::{BytesStart, Event};

impl PptxParser {
    pub(super) fn parse_shapes_from_xml(
        &mut self,
        xml: &str,
        slide_path: &str,
        relationships: &Relationships,
        zip: &mut impl PackageReader,
    ) -> Result<Vec<NodeId>, ParseError> {
        let mut reader = Reader::from_str(xml);
        let config = reader.config_mut();
        config.trim_text(true);
        config.check_end_names = true;
        let mut buf = Vec::new();
        let mut shapes = Vec::new();
        let mut root_name: Option<Vec<u8>> = None;
        let mut root_closed = false;

        loop {
            let event = read_event(&mut reader, &mut buf, slide_path)?;
            if root_closed && matches!(&event, Event::Start(_) | Event::Empty(_)) {
                return Err(xml_error(
                    slide_path,
                    "XML document contains multiple roots",
                ));
            }
            match &event {
                Event::Start(e) if root_name.is_none() => {
                    root_name = Some(e.name().as_ref().to_vec());
                }
                Event::Empty(e) if root_name.is_none() => {
                    root_name = Some(e.name().as_ref().to_vec());
                    root_closed = true;
                }
                Event::End(e)
                    if root_name
                        .as_deref()
                        .is_some_and(|name| name == e.name().as_ref()) =>
                {
                    root_closed = true;
                }
                Event::Eof if root_closed => break,
                Event::Eof => {
                    return Err(xml_error(
                        slide_path,
                        "XML document ends before its root is closed",
                    ));
                }
                _ => {}
            }
            if let Event::Start(e) = event {
                match local_name(e.name().as_ref()) {
                    b"sp" => {
                        let shape =
                            self.parse_shape_sp(&mut reader, &e, slide_path, relationships)?;
                        let id = shape.id;
                        self.store.insert(IRNode::Shape(shape));
                        shapes.push(id);
                    }
                    b"pic" => {
                        let shape =
                            self.parse_shape_pic(&mut reader, &e, slide_path, relationships)?;
                        let id = shape.id;
                        self.store.insert(IRNode::Shape(shape));
                        shapes.push(id);
                    }
                    b"graphicFrame" => {
                        let shape = self.parse_shape_graphic_frame(
                            &mut reader,
                            &e,
                            slide_path,
                            relationships,
                            zip,
                        )?;
                        let id = shape.id;
                        self.store.insert(IRNode::Shape(shape));
                        shapes.push(id);
                    }
                    b"grpSp" => {
                        let shape =
                            self.parse_shape_group(&mut reader, &e, slide_path, relationships)?;
                        let id = shape.id;
                        self.store.insert(IRNode::Shape(shape));
                        shapes.push(id);
                    }
                    _ => {}
                }
            }
            buf.clear();
        }

        Ok(shapes)
    }

    pub(super) fn parse_shape_sp(
        &mut self,
        reader: &mut Reader<&[u8]>,
        _start: &BytesStart,
        slide_path: &str,
        relationships: &Relationships,
    ) -> Result<Shape, ParseError> {
        let mut shape = Shape::new(ShapeType::Unknown);
        shape.span = Some(SourceSpan::new(slide_path));

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                    b"cNvPr" => {
                        parse_shape_non_visual_props(&e, &mut shape, slide_path)?;
                    }
                    b"hlinkClick" => {
                        self.attach_hyperlink(&mut shape, &e, relationships, slide_path)?;
                    }
                    b"spPr" => {
                        parse_shape_properties(reader, &mut shape, slide_path)?;
                    }
                    b"txBody" => {
                        let text = parse_text_body(reader, slide_path)?;
                        shape.text = Some(text);
                        if matches!(shape.shape_type, ShapeType::Unknown) {
                            shape.shape_type = ShapeType::TextBox;
                        }
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                    b"cNvPr" => {
                        parse_shape_non_visual_props(&e, &mut shape, slide_path)?;
                    }
                    b"hlinkClick" => {
                        self.attach_hyperlink(&mut shape, &e, relationships, slide_path)?;
                    }
                    b"spPr" => {}
                    _ => {}
                },
                Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"sp" => {
                    break;
                }
                Ok(Event::Eof) => {
                    return Err(xml_error(slide_path, "unexpected EOF in shape XML"));
                }
                Err(e) => {
                    return Err(xml_error(slide_path, e));
                }
                _ => {}
            }
            buf.clear();
        }

        if matches!(shape.shape_type, ShapeType::Unknown) {
            shape.shape_type = ShapeType::Rectangle;
        }

        Ok(shape)
    }

    pub(super) fn parse_shape_group(
        &mut self,
        reader: &mut Reader<&[u8]>,
        _start: &BytesStart,
        slide_path: &str,
        _relationships: &Relationships,
    ) -> Result<Shape, ParseError> {
        let mut shape = Shape::new(ShapeType::Group);
        shape.span = Some(SourceSpan::new(slide_path));

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                    b"cNvPr" => {
                        parse_non_visual_name(&e, &mut shape, slide_path)?;
                    }
                    b"grpSpPr" => {
                        parse_group_properties(reader, &mut shape, slide_path)?;
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                    b"cNvPr" => {
                        parse_non_visual_name(&e, &mut shape, slide_path)?;
                    }
                    b"grpSpPr" => {}
                    _ => {}
                },
                Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"grpSp" => {
                    break;
                }
                Ok(Event::Eof) => {
                    return Err(xml_error(slide_path, "unexpected EOF in group shape XML"));
                }
                Err(e) => {
                    return Err(xml_error(slide_path, e));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(shape)
    }

    pub(super) fn attach_hyperlink(
        &mut self,
        shape: &mut Shape,
        element: &BytesStart,
        relationships: &Relationships,
        slide_path: &str,
    ) -> Result<(), ParseError> {
        let Some(rel_id) = try_attr_value_by_suffix(element, &[b":id"], slide_path)? else {
            return Ok(());
        };
        let Some(rel) = relationships.get(&rel_id) else {
            return Ok(());
        };

        shape.hyperlink = Some(rel.target.clone());

        if rel.target_mode == TargetMode::External {
            let ref_type = classify_relationship(&rel.rel_type);
            self.add_external_reference(rel, ref_type, slide_path);
        }

        Ok(())
    }
}

fn parse_shape_non_visual_props(
    start: &BytesStart<'_>,
    shape: &mut Shape,
    slide_path: &str,
) -> Result<(), ParseError> {
    if let Some(name) = try_attr_value(start, b"name", slide_path)? {
        shape.name = Some(name);
    }
    if let Some(alt_text) = try_attr_value(start, b"descr", slide_path)? {
        shape.alt_text = Some(alt_text);
    }
    Ok(())
}

fn parse_non_visual_name(
    start: &BytesStart<'_>,
    shape: &mut Shape,
    slide_path: &str,
) -> Result<(), ParseError> {
    if let Some(name) = try_attr_value(start, b"name", slide_path)? {
        shape.name = Some(name);
    }
    Ok(())
}

fn parse_group_properties(
    reader: &mut Reader<&[u8]>,
    shape: &mut Shape,
    slide_path: &str,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"xfrm" => {
                parse_transform(reader, &mut shape.transform, slide_path)?;
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"grpSpPr" => {
                break;
            }
            Ok(Event::Eof) => {
                return Err(xml_error(
                    slide_path,
                    "unexpected EOF in group shape properties XML",
                ));
            }
            Err(e) => {
                return Err(xml_error(slide_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(())
}
