use super::graphic_frame::GraphicFrameState;
use super::{
    extract_c_sld_name, parse_comments, parse_shape_properties, parse_slide_layout_meta, PptxParser,
};
use crate::error::ParseError;
use crate::ooxml::relationships::{rel_type, Relationships, TargetMode};
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::reader_from_str;
use crate::zip_handler::PackageReader;
use docir_core::ir::{IRNode, Shape, ShapeType, Slide, SlideAnimation, SlideTransition};
use docir_core::security::ExternalRefType;
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

impl PptxParser {
    pub(super) fn parse_slide(
        &mut self,
        zip: &mut impl PackageReader,
        xml: &str,
        slide_number: u32,
        slide_path: &str,
        relationships: &Relationships,
        notes: (Option<&str>, Option<NodeId>),
    ) -> Result<NodeId, ParseError> {
        let (notes_text, notes_slide_id) = notes;
        let mut slide = self.build_slide_shell(slide_number, slide_path, relationships);

        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => self.handle_slide_start_event(
                    &mut reader,
                    &e,
                    slide_path,
                    relationships,
                    zip,
                    &mut slide,
                )?,
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: slide_path.to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        self.attach_slide_notes(&mut slide, notes_text, notes_slide_id);
        self.attach_slide_comments(zip, slide_path, relationships, &mut slide)?;

        let slide_id = slide.id;
        self.store.insert(IRNode::Slide(slide));
        Ok(slide_id)
    }

    fn build_slide_shell(
        &self,
        slide_number: u32,
        slide_path: &str,
        relationships: &Relationships,
    ) -> Slide {
        let mut slide = Slide::new(slide_number);
        slide.span = Some(SourceSpan::new(slide_path));
        if let Some(rel) = relationships.get_first_by_type(rel_type::SLIDE_LAYOUT) {
            slide.layout_id = Some(Relationships::resolve_target(slide_path, &rel.target));
        }
        if let Some(rel) = relationships.get_first_by_type(rel_type::SLIDE_MASTER) {
            slide.master_id = Some(Relationships::resolve_target(slide_path, &rel.target));
        }
        slide
    }

    fn handle_slide_start_event(
        &mut self,
        reader: &mut Reader<&[u8]>,
        event: &BytesStart<'_>,
        slide_path: &str,
        relationships: &Relationships,
        zip: &mut impl PackageReader,
        slide: &mut Slide,
    ) -> Result<(), ParseError> {
        match event.name().as_ref() {
            b"p:sld" => update_slide_visibility(slide, event),
            b"p:cSld" => update_slide_name(slide, event),
            b"p:sp" => {
                let shape = self.parse_shape_sp(reader, event, slide_path, relationships)?;
                self.push_slide_shape(slide, shape);
            }
            b"p:pic" => {
                let shape = self.parse_shape_pic(reader, event, slide_path, relationships)?;
                self.push_slide_shape(slide, shape);
            }
            b"p:graphicFrame" => {
                let shape =
                    self.parse_shape_graphic_frame(reader, event, slide_path, relationships, zip)?;
                self.push_slide_shape(slide, shape);
            }
            b"p:grpSp" => {
                let shape = self.parse_shape_group(reader, event, slide_path, relationships)?;
                self.push_slide_shape(slide, shape);
            }
            b"p:transition" => {
                slide.transition = Some(Self::parse_slide_transition(reader, event, slide_path)?)
            }
            b"p:timing" => {
                slide.animations = Self::parse_slide_animations(reader, slide_path, relationships)?
            }
            _ => {}
        }
        Ok(())
    }

    fn push_slide_shape(&mut self, slide: &mut Slide, shape: Shape) {
        let id = shape.id;
        self.store.insert(IRNode::Shape(shape));
        slide.shapes.push(id);
    }

    fn attach_slide_notes(
        &self,
        slide: &mut Slide,
        notes_text: Option<&str>,
        notes_slide_id: Option<NodeId>,
    ) {
        if let Some(notes) = notes_text {
            if !notes.trim().is_empty() {
                slide.notes = Some(notes.to_string());
            }
        }
        slide.notes_slide = notes_slide_id;
    }

    fn attach_slide_comments(
        &mut self,
        zip: &mut impl PackageReader,
        slide_path: &str,
        relationships: &Relationships,
        slide: &mut Slide,
    ) -> Result<(), ParseError> {
        let Some(rel) = relationships
            .by_id
            .values()
            .find(|r| r.rel_type.contains("comments"))
        else {
            return Ok(());
        };

        let comments_path = Relationships::resolve_target(slide_path, &rel.target);
        if !zip.contains(&comments_path) {
            return Ok(());
        }

        let comments_xml = zip.read_file_string(&comments_path)?;
        let comments = parse_comments(&comments_xml, &comments_path, &self.comment_authors)?;
        for comment in comments {
            let id = comment.id;
            self.store.insert(IRNode::PptxComment(comment));
            slide.comments.push(id);
        }
        Ok(())
    }

    pub(super) fn parse_slide_layout(
        &mut self,
        xml: &str,
        layout_path: &str,
        relationships: &Relationships,
        zip: &mut impl PackageReader,
    ) -> Result<NodeId, ParseError> {
        let mut layout = docir_core::ir::SlideLayout::new();
        layout.span = Some(SourceSpan::new(layout_path));
        layout.name = extract_c_sld_name(xml);
        let meta = parse_slide_layout_meta(xml, layout_path)?;
        layout.layout_type = meta.layout_type;
        layout.matching_name = meta.matching_name;
        layout.preserve = meta.preserve;
        layout.show_master_sp = meta.show_master_sp;
        layout.show_master_ph_anim = meta.show_master_ph_anim;
        layout.shapes = self.parse_shapes_from_xml(xml, layout_path, relationships, zip)?;
        let id = layout.id;
        self.store.insert(IRNode::SlideLayout(layout));
        Ok(id)
    }

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
                    );
                    if e.name().as_ref() == b"p:spPr" {
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
                    );
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"p:pic" {
                        break;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: slide_path.to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        let primary_rel = embed_rel.clone().or(link_rel.clone());
        if let Some(rel_id) = primary_rel {
            if let Some(rel) = relationships.get(&rel_id) {
                shape.relationship_id = Some(rel_id.clone());
                let resolved = if rel.target_mode == TargetMode::External {
                    rel.target.clone()
                } else {
                    Relationships::resolve_target(slide_path, &rel.target)
                };
                shape.media_target = Some(resolved);
                if rel.rel_type.contains("audio") {
                    shape.shape_type = ShapeType::Audio;
                } else if rel.rel_type.contains("video") {
                    shape.shape_type = ShapeType::Video;
                }
                if rel.target_mode == TargetMode::External {
                    let ref_type =
                        if rel.rel_type.contains("audio") || rel.rel_type.contains("video") {
                            ExternalRefType::Other
                        } else {
                            ExternalRefType::Image
                        };
                    self.add_external_reference(rel, ref_type, slide_path);
                }
            }
        }

        if let (Some(link_id), Some(embed_id)) = (link_rel.clone(), embed_rel.clone()) {
            if link_id != embed_id {
                if let Some(rel) = relationships.get(&link_id) {
                    if rel.target_mode == TargetMode::External {
                        let ref_type =
                            if rel.rel_type.contains("audio") || rel.rel_type.contains("video") {
                                ExternalRefType::Other
                            } else {
                                ExternalRefType::Image
                            };
                        self.add_external_reference(rel, ref_type, slide_path);
                    }
                }
            }
        }

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
    ) {
        match event.name().as_ref() {
            b"p:cNvPr" => {
                for attr in event.attributes().flatten() {
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
            b"a:hlinkClick" => {
                self.attach_hyperlink(shape, event, relationships, slide_path);
            }
            b"a:blip" => {
                for attr in event.attributes().flatten() {
                    if attr.key.as_ref() == b"r:embed" {
                        *embed_rel = Some(lossy_attr_value(&attr).to_string());
                    } else if attr.key.as_ref() == b"r:link" {
                        *link_rel = Some(lossy_attr_value(&attr).to_string());
                    }
                }
            }
            _ => {}
        }
    }

    pub(super) fn parse_shape_graphic_frame(
        &mut self,
        reader: &mut Reader<&[u8]>,
        _start: &BytesStart,
        slide_path: &str,
        relationships: &Relationships,
        zip: &mut impl PackageReader,
    ) -> Result<Shape, ParseError> {
        let mut shape = Shape::new(ShapeType::Custom);
        shape.span = Some(SourceSpan::new(slide_path));

        let mut buf = Vec::new();
        let mut state = GraphicFrameState::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    self.handle_graphic_frame_start(
                        &e,
                        reader,
                        slide_path,
                        relationships,
                        &mut shape,
                        &mut state,
                    )?;
                }
                Ok(Event::Empty(e)) => {
                    self.handle_graphic_frame_empty(
                        &e,
                        reader,
                        slide_path,
                        relationships,
                        &mut shape,
                        &mut state,
                    )?;
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"p:graphicFrame" {
                        break;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: slide_path.to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        self.apply_graphic_frame_relationships(&mut shape, slide_path, relationships, zip, &state)?;

        shape.table = state.table_id;

        Ok(shape)
    }

    fn parse_slide_transition(
        reader: &mut Reader<&[u8]>,
        start: &BytesStart,
        slide_path: &str,
    ) -> Result<SlideTransition, ParseError> {
        let mut transition = SlideTransition {
            transition_type: None,
            speed: None,
            advance_on_click: None,
            advance_after_ms: None,
            duration_ms: None,
        };

        for attr in start.attributes().flatten() {
            match attr.key.as_ref() {
                b"spd" => transition.speed = Some(lossy_attr_value(&attr).to_string()),
                b"advClick" => {
                    let v = lossy_attr_value(&attr);
                    transition.advance_on_click = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                }
                b"advTm" => {
                    transition.advance_after_ms = lossy_attr_value(&attr).parse::<u32>().ok();
                }
                b"dur" => {
                    transition.duration_ms = lossy_attr_value(&attr).parse::<u32>().ok();
                }
                _ => {}
            }
        }

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if transition.transition_type.is_none() {
                        transition.transition_type =
                            Some(String::from_utf8_lossy(e.name().as_ref()).to_string());
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"p:transition" {
                        break;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: slide_path.to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(transition)
    }

    fn parse_slide_animations(
        reader: &mut Reader<&[u8]>,
        slide_path: &str,
        relationships: &Relationships,
    ) -> Result<Vec<SlideAnimation>, ParseError> {
        let mut animations: Vec<SlideAnimation> = Vec::new();
        let mut current_index: Option<usize> = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let name = e.name().as_ref().to_vec();
                    if is_standard_animation(&name) || is_media_animation(&name) {
                        let anim = build_animation_from_event(
                            &name,
                            e.attributes().flatten(),
                            slide_path,
                            relationships,
                        );
                        animations.push(anim);
                        current_index = Some(animations.len() - 1);
                    } else if name.as_slice() == b"p:spTgt" {
                        apply_sp_target(&mut animations, current_index, e.attributes().flatten());
                    }
                }
                Ok(Event::Empty(e)) => {
                    let name = e.name().as_ref().to_vec();
                    if name.as_slice() == b"p:spTgt" {
                        apply_sp_target(&mut animations, current_index, e.attributes().flatten());
                    } else if is_media_animation(&name) {
                        let anim = build_animation_from_event(
                            &name,
                            e.attributes().flatten(),
                            slide_path,
                            relationships,
                        );
                        animations.push(anim);
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"p:timing" {
                        break;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: slide_path.to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(animations)
    }
}

fn update_slide_visibility(slide: &mut Slide, event: &BytesStart<'_>) {
    for attr in event.attributes().flatten() {
        if attr.key.as_ref() == b"show" {
            let v = lossy_attr_value(&attr);
            if v == "0" || v.eq_ignore_ascii_case("false") {
                slide.hidden = true;
            }
        }
    }
}

fn update_slide_name(slide: &mut Slide, event: &BytesStart<'_>) {
    for attr in event.attributes().flatten() {
        if attr.key.as_ref() == b"name" {
            slide.name = Some(lossy_attr_value(&attr).to_string());
        }
    }
}

fn is_standard_animation(name: &[u8]) -> bool {
    matches!(
        name,
        b"p:anim" | b"p:animEffect" | b"p:animMotion" | b"p:animRot" | b"p:animScale" | b"p:seq"
    )
}

fn is_media_animation(name: &[u8]) -> bool {
    name.ends_with(b":audio") || name.ends_with(b":video")
}

fn build_animation_from_event<'a, I>(
    name: &[u8],
    attrs: I,
    slide_path: &str,
    relationships: &Relationships,
) -> SlideAnimation
where
    I: Iterator<Item = quick_xml::events::attributes::Attribute<'a>>,
{
    let mut anim = SlideAnimation {
        animation_type: String::from_utf8_lossy(name).to_string(),
        target: None,
        duration_ms: None,
        preset_id: None,
        preset_class: None,
        media_asset: None,
    };

    for attr in attrs {
        match attr.key.as_ref() {
            b"dur" => {
                anim.duration_ms = lossy_attr_value(&attr).parse::<u32>().ok();
            }
            b"presetID" => {
                anim.preset_id = Some(lossy_attr_value(&attr).to_string());
            }
            b"presetClass" => {
                anim.preset_class = Some(lossy_attr_value(&attr).to_string());
            }
            b"r:link" | b"r:embed" => {
                let rel_id = lossy_attr_value(&attr).to_string();
                anim.target = Some(resolve_animation_target(slide_path, relationships, rel_id));
            }
            _ => {}
        }
    }

    anim
}

fn resolve_animation_target(
    slide_path: &str,
    relationships: &Relationships,
    rel_id: String,
) -> String {
    if let Some(rel) = relationships.get(&rel_id) {
        Relationships::resolve_target(slide_path, &rel.target)
    } else {
        rel_id
    }
}

fn apply_sp_target<'a, I>(animations: &mut [SlideAnimation], current_index: Option<usize>, attrs: I)
where
    I: Iterator<Item = quick_xml::events::attributes::Attribute<'a>>,
{
    let Some(idx) = current_index else {
        return;
    };
    for attr in attrs {
        if attr.key.as_ref() == b"spid" {
            animations[idx].target = Some(lossy_attr_value(&attr).to_string());
        }
    }
}
