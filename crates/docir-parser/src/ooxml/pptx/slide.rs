use super::graphic_frame::GraphicFrameState;
use super::{PptxParser, extract_c_sld_name, parse_comments, parse_slide_layout_meta};
use crate::error::ParseError;
use crate::ooxml::relationships::{Relationships, TargetMode, rel_type};
use crate::xml_utils::local_name;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::reader_from_str;
use crate::xml_utils::xml_error;
use crate::zip_handler::PackageReader;
use docir_core::ir::{IRNode, Shape, ShapeType, Slide, SlideAnimation, SlideTransition};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

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
                    return Err(xml_error(slide_path, e));
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
        match local_name(event.name().as_ref()) {
            b"sld" => update_slide_visibility(slide, event, slide_path)?,
            b"cSld" => update_slide_name(slide, event, slide_path)?,
            b"sp" => {
                let shape = self.parse_shape_sp(reader, event, slide_path, relationships)?;
                self.push_slide_shape(slide, shape);
            }
            b"pic" => {
                let shape = self.parse_shape_pic(reader, event, slide_path, relationships)?;
                self.push_slide_shape(slide, shape);
            }
            b"graphicFrame" => {
                let shape =
                    self.parse_shape_graphic_frame(reader, event, slide_path, relationships, zip)?;
                self.push_slide_shape(slide, shape);
            }
            b"grpSp" => {
                let shape = self.parse_shape_group(reader, event, slide_path, relationships)?;
                self.push_slide_shape(slide, shape);
            }
            b"transition" => {
                slide.transition = Some(Self::parse_slide_transition(reader, event, slide_path)?)
            }
            b"timing" => {
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
        if let Some(notes) = notes_text
            && !notes.trim().is_empty()
        {
            slide.notes = Some(notes.to_string());
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
                Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"graphicFrame" => {
                    break;
                }
                Ok(Event::Eof) => {
                    return Err(xml_error(slide_path, "unexpected EOF in graphic frame XML"));
                }
                Err(e) => {
                    return Err(xml_error(slide_path, e));
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

        for attr in start.attributes() {
            let attr = attr.map_err(|err| xml_error(slide_path, err))?;
            match attr.key.as_ref() {
                b"spd" => transition.speed = Some(lossy_attr_value(&attr).to_string()),
                b"advClick" => {
                    let value = lossy_attr_value(&attr);
                    transition.advance_on_click =
                        Some(value == "1" || value.eq_ignore_ascii_case("true"));
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
                Ok(Event::Start(e)) if transition.transition_type.is_none() => {
                    transition.transition_type =
                        Some(String::from_utf8_lossy(e.name().as_ref()).to_string());
                }
                Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"transition" => {
                    break;
                }
                Ok(Event::Eof) => {
                    return Err(xml_error(
                        slide_path,
                        "unexpected EOF in slide transition XML",
                    ));
                }
                Err(e) => {
                    return Err(xml_error(slide_path, e));
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
                    let local = local_name(&name);
                    if is_standard_animation(local) || is_media_animation(local) {
                        let anim =
                            build_animation_from_event(&name, &e, slide_path, relationships)?;
                        animations.push(anim);
                        current_index = Some(animations.len() - 1);
                    } else if local == b"spTgt" {
                        apply_sp_target(&mut animations, current_index, &e, slide_path)?;
                    }
                }
                Ok(Event::Empty(e)) => {
                    let name = e.name().as_ref().to_vec();
                    let local = local_name(&name);
                    if local == b"spTgt" {
                        apply_sp_target(&mut animations, current_index, &e, slide_path)?;
                    } else if is_media_animation(local) {
                        let anim =
                            build_animation_from_event(&name, &e, slide_path, relationships)?;
                        animations.push(anim);
                    }
                }
                Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"timing" => {
                    break;
                }
                Ok(Event::Eof) => {
                    return Err(xml_error(
                        slide_path,
                        "unexpected EOF in slide animation XML",
                    ));
                }
                Err(e) => {
                    return Err(xml_error(slide_path, e));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(animations)
    }
}

fn update_slide_visibility(
    slide: &mut Slide,
    event: &BytesStart<'_>,
    slide_path: &str,
) -> Result<(), ParseError> {
    for attr in event.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
        if attr.key.as_ref() == b"show" {
            let value = lossy_attr_value(&attr);
            if value == "0" || value.eq_ignore_ascii_case("false") {
                slide.hidden = true;
            }
        }
    }
    Ok(())
}

fn update_slide_name(
    slide: &mut Slide,
    event: &BytesStart<'_>,
    slide_path: &str,
) -> Result<(), ParseError> {
    for attr in event.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
        if attr.key.as_ref() == b"name" {
            slide.name = Some(lossy_attr_value(&attr).to_string());
        }
    }
    Ok(())
}

fn is_standard_animation(name: &[u8]) -> bool {
    matches!(
        name,
        b"anim" | b"animEffect" | b"animMotion" | b"animRot" | b"animScale" | b"seq"
    )
}

fn is_media_animation(name: &[u8]) -> bool {
    name == b"audio" || name == b"video"
}

fn build_animation_from_event(
    name: &[u8],
    event: &BytesStart<'_>,
    slide_path: &str,
    relationships: &Relationships,
) -> Result<SlideAnimation, ParseError> {
    let mut anim = SlideAnimation {
        animation_type: String::from_utf8_lossy(name).to_string(),
        target: None,
        duration_ms: None,
        preset_id: None,
        preset_class: None,
        media_asset: None,
    };

    for attr in event.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
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
            key if matches!(local_name(key), b"link" | b"embed") => {
                let rel_id = lossy_attr_value(&attr).to_string();
                anim.target = Some(resolve_animation_target(slide_path, relationships, rel_id));
            }
            _ => {}
        }
    }

    Ok(anim)
}

fn resolve_animation_target(
    slide_path: &str,
    relationships: &Relationships,
    rel_id: String,
) -> String {
    if let Some(rel) = relationships.get(&rel_id) {
        if rel.target_mode == TargetMode::External {
            rel.target.clone()
        } else {
            Relationships::resolve_target(slide_path, &rel.target)
        }
    } else {
        rel_id
    }
}

fn apply_sp_target(
    animations: &mut [SlideAnimation],
    current_index: Option<usize>,
    event: &BytesStart<'_>,
    slide_path: &str,
) -> Result<(), ParseError> {
    let mut target = None;
    for attr in event.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
        if attr.key.as_ref() == b"spid" {
            target = Some(lossy_attr_value(&attr).to_string());
        }
    }
    if let Some(idx) = current_index
        && let Some(target) = target
    {
        animations[idx].target = Some(target);
    }
    Ok(())
}
