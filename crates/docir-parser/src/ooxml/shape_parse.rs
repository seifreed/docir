use super::{
    attr_value, classify_relationship, parse_shape_properties, parse_text_body, parse_transform,
    read_event, PackageReader, ParseError, PptxParser, Reader, Relationships, Shape, ShapeType,
    SourceSpan, TargetMode,
};
use crate::xml_utils::{attr_value_by_suffix, local_name, xml_error};
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
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut shapes = Vec::new();

        loop {
            match read_event(&mut reader, &mut buf, slide_path)? {
                Event::Start(e) => match local_name(e.name().as_ref()) {
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
                },
                Event::Eof => break,
                _ => {}
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
                        parse_shape_non_visual_props(&e, &mut shape);
                    }
                    b"hlinkClick" => {
                        self.attach_hyperlink(&mut shape, &e, relationships, slide_path);
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
                        parse_shape_non_visual_props(&e, &mut shape);
                    }
                    b"hlinkClick" => {
                        self.attach_hyperlink(&mut shape, &e, relationships, slide_path);
                    }
                    b"spPr" => {}
                    _ => {}
                },
                Ok(Event::End(e)) => {
                    if local_name(e.name().as_ref()) == b"sp" {
                        break;
                    }
                }
                Ok(Event::Eof) => break,
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
                        parse_non_visual_name(&e, &mut shape);
                    }
                    b"grpSpPr" => {
                        parse_group_properties(reader, &mut shape, slide_path)?;
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                    b"cNvPr" => {
                        parse_non_visual_name(&e, &mut shape);
                    }
                    b"grpSpPr" => {}
                    _ => {}
                },
                Ok(Event::End(e)) => {
                    if local_name(e.name().as_ref()) == b"grpSp" {
                        break;
                    }
                }
                Ok(Event::Eof) => break,
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
    ) {
        let Some(rel_id) = attr_value_by_suffix(element, &[b":id"]) else {
            return;
        };
        let Some(rel) = relationships.get(&rel_id) else {
            return;
        };

        shape.hyperlink = Some(rel.target.clone());

        if rel.target_mode == TargetMode::External {
            let ref_type = classify_relationship(&rel.rel_type);
            self.add_external_reference(rel, ref_type, slide_path);
        }
    }
}

fn parse_shape_non_visual_props(start: &BytesStart<'_>, shape: &mut Shape) {
    if let Some(name) = attr_value(start, b"name") {
        shape.name = Some(name);
    }
    if let Some(alt_text) = attr_value(start, b"descr") {
        shape.alt_text = Some(alt_text);
    }
}

fn parse_non_visual_name(start: &BytesStart<'_>, shape: &mut Shape) {
    if let Some(name) = attr_value(start, b"name") {
        shape.name = Some(name);
    }
}

fn parse_group_properties(
    reader: &mut Reader<&[u8]>,
    shape: &mut Shape,
    slide_path: &str,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if local_name(e.name().as_ref()) == b"xfrm" {
                    parse_transform(reader, &mut shape.transform, slide_path)?;
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"grpSpPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(slide_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(())
}
