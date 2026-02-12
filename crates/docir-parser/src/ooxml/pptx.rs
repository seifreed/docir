//! PPTX presentation and slide parsing.

use crate::diagnostics::push_warning;
use crate::error::ParseError;
use crate::ooxml::relationships::{rel_type, Relationship, Relationships, TargetMode};
use crate::zip_handler::SecureZipReader;
use docir_core::ir::{
    Diagnostics, Document, GridColumn, IRNode, NotesSlide, Paragraph, PptxComment,
    PptxCommentAuthor, PresentationInfo, PresentationProperties, PresentationTag, Run, Shape,
    ShapeText, ShapeTextParagraph, ShapeTextRun, ShapeTransform, ShapeType, Slide, SlideAnimation,
    SlideSize, SlideTransition, SmartArtPart, Table, TableCell, TableRow, TableStyle,
    TableStyleSet, TextAlignment, ViewProperties,
};
use docir_core::security::{ExternalRefType, ExternalReference, SecurityInfo};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

mod shapes;

use shapes::parse_shape_properties;

/// PPTX parser for presentation.xml and slides.
pub struct PptxParser {
    store: IrStore,
    security_info: SecurityInfo,
    external_rel_ids: HashSet<String>,
    chart_nodes: Vec<NodeId>,
    comment_authors: HashMap<u32, (Option<String>, Option<String>)>,
    diagnostics: Diagnostics,
}

impl PptxParser {
    /// Creates a new PPTX parser.
    pub fn new() -> Self {
        Self {
            store: IrStore::new(),
            security_info: SecurityInfo::default(),
            external_rel_ids: HashSet::new(),
            chart_nodes: Vec::new(),
            comment_authors: HashMap::new(),
            diagnostics: Diagnostics::new(),
        }
    }

    fn set_comment_authors(&mut self, authors: &[PptxCommentAuthor]) {
        self.comment_authors.clear();
        for author in authors {
            self.comment_authors.insert(
                author.author_id,
                (author.name.clone(), author.initials.clone()),
            );
        }
    }

    /// Parses the presentation and slides.
    pub fn parse_presentation<R: Read + Seek>(
        &mut self,
        zip: &mut SecureZipReader<R>,
        presentation_xml: &str,
        presentation_rels: &Relationships,
        presentation_path: &str,
    ) -> Result<NodeId, ParseError> {
        let mut document = Document::new(DocumentFormat::Presentation);
        document.span = Some(SourceSpan::new(presentation_path));

        self.process_external_relationships(presentation_rels, presentation_path);

        let slide_refs = parse_slide_list(presentation_xml)?;

        if let Some(mut info) = parse_presentation_info(presentation_xml, presentation_path)? {
            info.span = Some(SourceSpan::new(presentation_path));
            let id = info.id;
            self.store.insert(IRNode::PresentationInfo(info));
            document.shared_parts.push(id);
        }

        for (index, rel_id) in slide_refs.into_iter().enumerate() {
            let rel = match presentation_rels.get(&rel_id) {
                Some(rel) => rel,
                None => {
                    push_warning(
                        &mut self.diagnostics,
                        "MISSING_RELATIONSHIP",
                        format!("Missing relationship for slide relId {}", rel_id),
                        Some(presentation_path),
                    );
                    continue;
                }
            };
            let slide_path = Relationships::resolve_target(presentation_path, &rel.target);

            let slide_xml = zip.read_file_string(&slide_path)?;

            let rels_path = get_rels_path(&slide_path);
            let slide_rels = if zip.contains(&rels_path) {
                let rels_xml = zip.read_file_string(&rels_path)?;
                Relationships::parse(&rels_xml)?
            } else {
                Relationships::default()
            };

            self.process_external_relationships(&slide_rels, &slide_path);

            let mut notes_slide_id: Option<NodeId> = None;
            let notes_text = if let Some(rel) = slide_rels.get_first_by_type(rel_type::NOTES_SLIDE)
            {
                let notes_path = Relationships::resolve_target(&slide_path, &rel.target);
                if let Ok(notes_xml) = zip.read_file_string(&notes_path) {
                    let notes_rels = if zip.contains(&get_rels_path(&notes_path)) {
                        let rels_xml = zip.read_file_string(&get_rels_path(&notes_path))?;
                        Relationships::parse(&rels_xml)?
                    } else {
                        Relationships::default()
                    };
                    let (notes_node, notes_text) =
                        parse_notes_slide(&notes_xml, &notes_path, &notes_rels, self, zip)?;
                    let notes_id = notes_node.id;
                    self.store.insert(IRNode::NotesSlide(notes_node));
                    notes_slide_id = Some(notes_id);
                    Some(notes_text)
                } else {
                    None
                }
            } else {
                None
            };

            let slide_id = self.parse_slide(
                zip,
                &slide_xml,
                (index + 1) as u32,
                &slide_path,
                &slide_rels,
                notes_text.as_deref(),
                notes_slide_id,
            )?;
            document.content.push(slide_id);
        }

        // Presentation properties
        if zip.contains("ppt/presProps.xml") {
            let props_xml = zip.read_file_string("ppt/presProps.xml")?;
            let mut props = parse_presentation_properties(&props_xml, "ppt/presProps.xml")?;
            props.span = Some(SourceSpan::new("ppt/presProps.xml"));
            let id = props.id;
            self.store.insert(IRNode::PresentationProperties(props));
            document.shared_parts.push(id);
        }

        // View properties
        if zip.contains("ppt/viewProps.xml") {
            let view_xml = zip.read_file_string("ppt/viewProps.xml")?;
            let mut props = parse_view_properties(&view_xml, "ppt/viewProps.xml")?;
            props.span = Some(SourceSpan::new("ppt/viewProps.xml"));
            let id = props.id;
            self.store.insert(IRNode::ViewProperties(props));
            document.shared_parts.push(id);
        }

        // Table styles
        if zip.contains("ppt/tableStyles.xml") {
            let styles_xml = zip.read_file_string("ppt/tableStyles.xml")?;
            let mut styles = parse_table_styles(&styles_xml, "ppt/tableStyles.xml")?;
            styles.span = Some(SourceSpan::new("ppt/tableStyles.xml"));
            let id = styles.id;
            self.store.insert(IRNode::TableStyleSet(styles));
            document.shared_parts.push(id);
        }

        // Comment authors
        if zip.contains("ppt/commentAuthors.xml") {
            let authors_xml = zip.read_file_string("ppt/commentAuthors.xml")?;
            let authors = parse_comment_authors(&authors_xml, "ppt/commentAuthors.xml")?;
            self.set_comment_authors(&authors);
            for author in authors {
                let mut author = author;
                author.span = Some(SourceSpan::new("ppt/commentAuthors.xml"));
                let id = author.id;
                self.store.insert(IRNode::PptxCommentAuthor(author));
                document.shared_parts.push(id);
            }
        }

        // Tags
        let tag_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.starts_with("ppt/tags/") && p.ends_with(".xml"))
            .map(|s| s.to_string())
            .collect();
        for tag_path in tag_paths {
            let tag_xml = zip.read_file_string(&tag_path)?;
            let tags = parse_presentation_tags(&tag_xml, &tag_path)?;
            for tag in tags {
                let id = tag.id;
                self.store.insert(IRNode::PresentationTag(tag));
                document.shared_parts.push(id);
            }
        }

        // People part (coauthoring)
        if zip.contains("ppt/people.xml") {
            let xml = zip.read_file_string("ppt/people.xml")?;
            let mut people = crate::ooxml::shared::parse_people_part(&xml, "ppt/people.xml")?;
            people.span = Some(SourceSpan::new("ppt/people.xml"));
            let id = people.id;
            self.store.insert(IRNode::PeoplePart(people));
            document.shared_parts.push(id);
        }

        // SmartArt parts
        let diagram_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.starts_with("ppt/diagrams/") && p.ends_with(".xml"))
            .map(|s| s.to_string())
            .collect();
        for path in diagram_paths {
            let xml = zip.read_file_string(&path)?;
            let part = parse_smartart_part(&xml, &path)?;
            let id = part.id;
            self.store.insert(IRNode::SmartArtPart(part));
            document.shared_parts.push(id);
        }

        // Parse slide masters and layouts
        for rel in presentation_rels.get_by_type(rel_type::SLIDE_MASTER) {
            let master_path = Relationships::resolve_target(presentation_path, &rel.target);
            if !zip.contains(&master_path) {
                continue;
            }
            let master_xml = zip.read_file_string(&master_path)?;
            let rels_path = get_rels_path(&master_path);
            let master_rels = if zip.contains(&rels_path) {
                let rels_xml = zip.read_file_string(&rels_path)?;
                Relationships::parse(&rels_xml)?
            } else {
                Relationships::default()
            };

            let master_name = extract_c_sld_name(&master_xml);
            let master_shapes =
                self.parse_shapes_from_xml(&master_xml, &master_path, &master_rels, zip)?;
            let master_meta = parse_slide_master_meta(&master_xml, &master_path)?;

            let mut layout_ids = Vec::new();
            for layout_rel in master_rels.get_by_type(rel_type::SLIDE_LAYOUT) {
                let layout_path = Relationships::resolve_target(&master_path, &layout_rel.target);
                if !zip.contains(&layout_path) {
                    continue;
                }
                let layout_xml = zip.read_file_string(&layout_path)?;
                let layout_id =
                    self.parse_slide_layout(&layout_xml, &layout_path, &master_rels, zip)?;
                layout_ids.push(layout_id);
            }

            let mut master = docir_core::ir::SlideMaster::new();
            master.name = master_name;
            master.preserve = master_meta.preserve;
            master.show_master_sp = master_meta.show_master_sp;
            master.show_master_ph_anim = master_meta.show_master_ph_anim;
            master.shapes = master_shapes;
            master.layouts = layout_ids.clone();
            master.span = Some(SourceSpan::new(&master_path));
            let master_id = master.id;
            self.store.insert(IRNode::SlideMaster(master));

            document.shared_parts.push(master_id);
            document.shared_parts.extend(layout_ids);
        }

        // Parse notes master
        if let Some(rel) = presentation_rels.get_first_by_type(rel_type::NOTES_MASTER) {
            let notes_master_path = Relationships::resolve_target(presentation_path, &rel.target);
            if zip.contains(&notes_master_path) {
                let notes_master_xml = zip.read_file_string(&notes_master_path)?;
                let mut notes_master = docir_core::ir::NotesMaster::new();
                notes_master.name = extract_c_sld_name(&notes_master_xml);
                notes_master.shapes = self.parse_shapes_from_xml(
                    &notes_master_xml,
                    &notes_master_path,
                    presentation_rels,
                    zip,
                )?;
                notes_master.span = Some(SourceSpan::new(&notes_master_path));
                let id = notes_master.id;
                self.store.insert(IRNode::NotesMaster(notes_master));
                document.shared_parts.push(id);
            }
        }

        // Parse handout master
        if let Some(rel) = presentation_rels.get_first_by_type(rel_type::HANDOUT_MASTER) {
            let handout_path = Relationships::resolve_target(presentation_path, &rel.target);
            if zip.contains(&handout_path) {
                let handout_xml = zip.read_file_string(&handout_path)?;
                let mut handout = docir_core::ir::HandoutMaster::new();
                handout.name = extract_c_sld_name(&handout_xml);
                handout.shapes = self.parse_shapes_from_xml(
                    &handout_xml,
                    &handout_path,
                    presentation_rels,
                    zip,
                )?;
                handout.span = Some(SourceSpan::new(&handout_path));
                let id = handout.id;
                self.store.insert(IRNode::HandoutMaster(handout));
                document.shared_parts.push(id);
            }
        }

        document.shared_parts.extend(self.chart_nodes.drain(..));
        document.security = std::mem::take(&mut self.security_info);
        document.security.recalculate_threat_level();

        let mut diagnostics = std::mem::replace(&mut self.diagnostics, Diagnostics::new());
        if !diagnostics.entries.is_empty() {
            diagnostics.span = Some(SourceSpan::new(presentation_path));
            let diag_id = diagnostics.id;
            self.store.insert(IRNode::Diagnostics(diagnostics));
            document.diagnostics.push(diag_id);
        }

        let doc_id = document.id;
        self.store.insert(IRNode::Document(document));
        Ok(doc_id)
    }

    /// Returns the IR store.
    pub fn into_store(self) -> IrStore {
        self.store
    }

    fn parse_slide<R: Read + Seek>(
        &mut self,
        zip: &mut SecureZipReader<R>,
        xml: &str,
        slide_number: u32,
        slide_path: &str,
        relationships: &Relationships,
        notes_text: Option<&str>,
        notes_slide_id: Option<NodeId>,
    ) -> Result<NodeId, ParseError> {
        let mut slide = Slide::new(slide_number);
        slide.span = Some(SourceSpan::new(slide_path));

        if let Some(rel) = relationships.get_first_by_type(rel_type::SLIDE_LAYOUT) {
            slide.layout_id = Some(Relationships::resolve_target(slide_path, &rel.target));
        }
        if let Some(rel) = relationships.get_first_by_type(rel_type::SLIDE_MASTER) {
            slide.master_id = Some(Relationships::resolve_target(slide_path, &rel.target));
        }

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"p:sld" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"show" {
                                let v = String::from_utf8_lossy(&attr.value);
                                if v == "0" || v.eq_ignore_ascii_case("false") {
                                    slide.hidden = true;
                                }
                            }
                        }
                    }
                    b"p:cSld" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"name" {
                                slide.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    b"p:sp" => {
                        let shape =
                            self.parse_shape_sp(&mut reader, &e, slide_path, relationships)?;
                        let shape_id = shape.id;
                        self.store.insert(IRNode::Shape(shape));
                        slide.shapes.push(shape_id);
                    }
                    b"p:pic" => {
                        let shape =
                            self.parse_shape_pic(&mut reader, &e, slide_path, relationships)?;
                        let shape_id = shape.id;
                        self.store.insert(IRNode::Shape(shape));
                        slide.shapes.push(shape_id);
                    }
                    b"p:graphicFrame" => {
                        let shape = self.parse_shape_graphic_frame(
                            &mut reader,
                            &e,
                            slide_path,
                            relationships,
                            zip,
                        )?;
                        let shape_id = shape.id;
                        self.store.insert(IRNode::Shape(shape));
                        slide.shapes.push(shape_id);
                    }
                    b"p:grpSp" => {
                        let shape =
                            self.parse_shape_group(&mut reader, &e, slide_path, relationships)?;
                        let shape_id = shape.id;
                        self.store.insert(IRNode::Shape(shape));
                        slide.shapes.push(shape_id);
                    }
                    b"p:transition" => {
                        let transition = parse_slide_transition(&mut reader, &e, slide_path)?;
                        slide.transition = Some(transition);
                    }
                    b"p:timing" => {
                        let animations =
                            parse_slide_animations(&mut reader, slide_path, relationships)?;
                        slide.animations = animations;
                    }
                    _ => {}
                },
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

        if let Some(notes) = notes_text {
            if !notes.trim().is_empty() {
                slide.notes = Some(notes.to_string());
            }
        }
        slide.notes_slide = notes_slide_id;

        // Slide comments
        if let Some(rel) = relationships
            .by_id
            .values()
            .find(|r| r.rel_type.contains("comments"))
        {
            let comments_path = Relationships::resolve_target(slide_path, &rel.target);
            if zip.contains(&comments_path) {
                let comments_xml = zip.read_file_string(&comments_path)?;
                let comments =
                    parse_comments(&comments_xml, &comments_path, &self.comment_authors)?;
                for comment in comments {
                    let id = comment.id;
                    self.store.insert(IRNode::PptxComment(comment));
                    slide.comments.push(id);
                }
            }
        }

        let slide_id = slide.id;
        self.store.insert(IRNode::Slide(slide));
        Ok(slide_id)
    }

    fn parse_slide_layout<R: Read + Seek>(
        &mut self,
        xml: &str,
        layout_path: &str,
        relationships: &Relationships,
        zip: &mut SecureZipReader<R>,
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

    fn parse_shapes_from_xml<R: Read + Seek>(
        &mut self,
        xml: &str,
        slide_path: &str,
        relationships: &Relationships,
        zip: &mut SecureZipReader<R>,
    ) -> Result<Vec<NodeId>, ParseError> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut shapes = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"p:sp" => {
                        let shape =
                            self.parse_shape_sp(&mut reader, &e, slide_path, relationships)?;
                        let id = shape.id;
                        self.store.insert(IRNode::Shape(shape));
                        shapes.push(id);
                    }
                    b"p:pic" => {
                        let shape =
                            self.parse_shape_pic(&mut reader, &e, slide_path, relationships)?;
                        let id = shape.id;
                        self.store.insert(IRNode::Shape(shape));
                        shapes.push(id);
                    }
                    b"p:graphicFrame" => {
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
                    b"p:grpSp" => {
                        let shape =
                            self.parse_shape_group(&mut reader, &e, slide_path, relationships)?;
                        let id = shape.id;
                        self.store.insert(IRNode::Shape(shape));
                        shapes.push(id);
                    }
                    _ => {}
                },
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

        Ok(shapes)
    }

    fn parse_shape_sp(
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
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"p:cNvPr" => {
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
                    b"a:hlinkClick" => {
                        self.attach_hyperlink(&mut shape, &e, relationships, slide_path);
                    }
                    b"p:spPr" => {
                        parse_shape_properties(reader, &mut shape, slide_path)?;
                    }
                    b"p:txBody" => {
                        let text = parse_text_body(reader, slide_path)?;
                        shape.text = Some(text);
                        if matches!(shape.shape_type, ShapeType::Unknown) {
                            shape.shape_type = ShapeType::TextBox;
                        }
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => {
                    match e.name().as_ref() {
                        b"p:cNvPr" => {
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
                        b"a:hlinkClick" => {
                            self.attach_hyperlink(&mut shape, &e, relationships, slide_path);
                        }
                        b"p:spPr" => {
                            // No nested properties.
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"p:sp" {
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

        if matches!(shape.shape_type, ShapeType::Unknown) {
            shape.shape_type = ShapeType::Rectangle;
        }

        Ok(shape)
    }

    fn parse_shape_pic(
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
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"p:cNvPr" => {
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
                    b"a:hlinkClick" => {
                        self.attach_hyperlink(&mut shape, &e, relationships, slide_path);
                    }
                    b"a:blip" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r:embed" {
                                embed_rel = Some(String::from_utf8_lossy(&attr.value).to_string());
                            } else if attr.key.as_ref() == b"r:link" {
                                link_rel = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    b"p:spPr" => {
                        parse_shape_properties(reader, &mut shape, slide_path)?;
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => match e.name().as_ref() {
                    b"p:cNvPr" => {
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
                    b"a:hlinkClick" => {
                        self.attach_hyperlink(&mut shape, &e, relationships, slide_path);
                    }
                    b"a:blip" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r:embed" {
                                embed_rel = Some(String::from_utf8_lossy(&attr.value).to_string());
                            } else if attr.key.as_ref() == b"r:link" {
                                link_rel = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    b"p:spPr" => {}
                    _ => {}
                },
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

    fn parse_shape_graphic_frame<R: Read + Seek>(
        &mut self,
        reader: &mut Reader<&[u8]>,
        _start: &BytesStart,
        slide_path: &str,
        relationships: &Relationships,
        zip: &mut SecureZipReader<R>,
    ) -> Result<Shape, ParseError> {
        let mut shape = Shape::new(ShapeType::Custom);
        shape.span = Some(SourceSpan::new(slide_path));

        let mut buf = Vec::new();
        let mut chart_rel: Option<String> = None;
        let mut ole_rel: Option<String> = None;
        let mut table_id: Option<NodeId> = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"p:cNvPr" => {
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
                    b"a:hlinkClick" => {
                        self.attach_hyperlink(&mut shape, &e, relationships, slide_path);
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
                        table_id = Some(id);
                        shape.shape_type = ShapeType::Table;
                    }
                    _ if e.name().as_ref().ends_with(b"chart") => {
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            if chart_rel.is_none()
                                && (key == b"r:id" || key == b"id" || key.ends_with(b":id"))
                            {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val.starts_with("rId") {
                                    chart_rel = Some(val);
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
                                ole_rel = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        shape.shape_type = ShapeType::OleObject;
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => match e.name().as_ref() {
                    b"p:cNvPr" => {
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
                    b"a:hlinkClick" => {
                        self.attach_hyperlink(&mut shape, &e, relationships, slide_path);
                    }
                    b"p:xfrm" => {}
                    _ if e.name().as_ref().ends_with(b"graphicData") => {}
                    b"a:tbl" => {
                        let table = self.parse_pptx_table(reader, slide_path)?;
                        let id = table.id;
                        self.store.insert(IRNode::Table(table));
                        table_id = Some(id);
                        shape.shape_type = ShapeType::Table;
                    }
                    _ if e.name().as_ref().ends_with(b"chart") => {
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            if chart_rel.is_none()
                                && (key == b"r:id" || key == b"id" || key.ends_with(b":id"))
                            {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val.starts_with("rId") {
                                    chart_rel = Some(val);
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
                                ole_rel = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        shape.shape_type = ShapeType::OleObject;
                    }
                    _ => {}
                },
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

        if let Some(rel_id) = ole_rel {
            if let Some(rel) = relationships.get(&rel_id) {
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
                        relationship_id: Some(rel_id),
                        ..ext_ref
                    };
                    let ext_id = ext_ref.id;
                    self.store.insert(IRNode::ExternalReference(ext_ref));
                    self.security_info.external_refs.push(ext_id);
                }
            }
        } else if let Some(rel_id) = chart_rel {
            if let Some(rel) = relationships.get(&rel_id) {
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
                        relationship_id: Some(rel_id),
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

        shape.table = table_id;

        Ok(shape)
    }

    fn parse_shape_group(
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
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"p:cNvPr" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"name" {
                                shape.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    b"p:grpSpPr" => {
                        parse_group_properties(reader, &mut shape, slide_path)?;
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => match e.name().as_ref() {
                    b"p:cNvPr" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"name" {
                                shape.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    b"p:grpSpPr" => {}
                    _ => {}
                },
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"p:grpSp" {
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

        Ok(shape)
    }

    fn parse_pptx_table(
        &mut self,
        reader: &mut Reader<&[u8]>,
        slide_path: &str,
    ) -> Result<Table, ParseError> {
        let mut table = Table::new();
        table.span = Some(SourceSpan::new(slide_path));

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"a:gridCol" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w" {
                                if let Ok(width) =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    table.grid.push(GridColumn { width });
                                }
                            }
                        }
                    }
                    b"a:tr" => {
                        let row = self.parse_pptx_table_row(reader, slide_path)?;
                        let id = row.id;
                        self.store.insert(IRNode::TableRow(row));
                        table.rows.push(id);
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"a:gridCol" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w" {
                                if let Ok(width) =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    table.grid.push(GridColumn { width });
                                }
                            }
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"a:tbl" {
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

        Ok(table)
    }

    fn parse_pptx_table_row(
        &mut self,
        reader: &mut Reader<&[u8]>,
        slide_path: &str,
    ) -> Result<TableRow, ParseError> {
        let mut row = TableRow::new();
        row.span = Some(SourceSpan::new(slide_path));

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"a:tc" => {
                        let cell = self.parse_pptx_table_cell(reader, slide_path)?;
                        let id = cell.id;
                        self.store.insert(IRNode::TableCell(cell));
                        row.cells.push(id);
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"a:tr" {
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

        Ok(row)
    }

    fn parse_pptx_table_cell(
        &mut self,
        reader: &mut Reader<&[u8]>,
        slide_path: &str,
    ) -> Result<TableCell, ParseError> {
        let mut cell = TableCell::new();
        cell.span = Some(SourceSpan::new(slide_path));

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"a:txBody" => {
                        let text = parse_text_body_table(reader, slide_path)?;
                        let plain = shape_text_to_plain(&text);
                        if !plain.is_empty() {
                            let mut para = Paragraph::new();
                            let run = Run::new(plain);
                            let run_id = run.id;
                            self.store.insert(IRNode::Run(run));
                            para.runs.push(run_id);
                            let para_id = para.id;
                            self.store.insert(IRNode::Paragraph(para));
                            cell.content.push(para_id);
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"a:tc" {
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

        Ok(cell)
    }

    fn attach_hyperlink(
        &mut self,
        shape: &mut Shape,
        element: &BytesStart,
        relationships: &Relationships,
        slide_path: &str,
    ) {
        let mut rel_id = None;
        for attr in element.attributes().flatten() {
            if attr.key.as_ref() == b"r:id" {
                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
        }

        let Some(rel_id) = rel_id else {
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

    fn process_external_relationships(&mut self, rels: &Relationships, file_path: &str) {
        for rel in rels.external_relationships() {
            let ref_type = classify_relationship(&rel.rel_type);
            self.add_external_reference(rel, ref_type, file_path);
        }
    }

    fn add_external_reference(
        &mut self,
        rel: &Relationship,
        ref_type: ExternalRefType,
        file_path: &str,
    ) {
        let key = format!("{file_path}::{id}", id = rel.id);
        if !self.external_rel_ids.insert(key) {
            return;
        }

        let mut ext_ref = ExternalReference::new(ref_type, &rel.target);
        ext_ref.relationship_id = Some(rel.id.clone());
        ext_ref.relationship_type = Some(rel.rel_type.clone());
        ext_ref.span = Some(SourceSpan::new(file_path).with_relationship(rel.id.clone()));

        let ext_id = ext_ref.id;
        self.store.insert(IRNode::ExternalReference(ext_ref));
        self.security_info.external_refs.push(ext_id);
    }
}

impl Default for PptxParser {
    fn default() -> Self {
        Self::new()
    }
}

fn parse_slide_list(xml: &str) -> Result<Vec<String>, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut slide_ids = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"p:sldId" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"r:id" {
                            slide_ids.push(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "ppt/presentation.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(slide_ids)
}

fn parse_presentation_info(xml: &str, path: &str) -> Result<Option<PresentationInfo>, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut info = PresentationInfo::new();
    let mut buf = Vec::new();
    let mut found = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_bytes = e.name().as_ref().to_vec();
                let name = name_bytes.as_slice();
                if name == b"p:sldSz" {
                    let mut cx = None;
                    let mut cy = None;
                    let mut size_type = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"cx" => cx = String::from_utf8_lossy(&attr.value).parse::<u64>().ok(),
                            b"cy" => cy = String::from_utf8_lossy(&attr.value).parse::<u64>().ok(),
                            b"type" => {
                                size_type = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            _ => {}
                        }
                    }
                    if let (Some(cx), Some(cy)) = (cx, cy) {
                        info.slide_size = Some(SlideSize { cx, cy, size_type });
                        found = true;
                    }
                } else if name == b"p:notesSz" {
                    let mut cx = None;
                    let mut cy = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"cx" => cx = String::from_utf8_lossy(&attr.value).parse::<u64>().ok(),
                            b"cy" => cy = String::from_utf8_lossy(&attr.value).parse::<u64>().ok(),
                            _ => {}
                        }
                    }
                    if let (Some(cx), Some(cy)) = (cx, cy) {
                        info.notes_size = Some(SlideSize {
                            cx,
                            cy,
                            size_type: None,
                        });
                        found = true;
                    }
                } else if name == b"p:showPr" {
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let val = String::from_utf8_lossy(&attr.value);
                        match key {
                            b"showType" => info.show_type = Some(val.to_string()),
                            b"loop" => {
                                info.show_loop =
                                    Some(val == "1" || val.eq_ignore_ascii_case("true"))
                            }
                            b"showNarration" => {
                                info.show_narration =
                                    Some(val == "1" || val.eq_ignore_ascii_case("true"))
                            }
                            b"showAnimation" => {
                                info.show_animation =
                                    Some(val == "1" || val.eq_ignore_ascii_case("true"))
                            }
                            b"useTimings" => {
                                info.use_timings =
                                    Some(val == "1" || val.eq_ignore_ascii_case("true"))
                            }
                            _ => {}
                        }
                    }
                    found = true;
                } else if name == b"p:presentation" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"firstSlideNum" {
                            info.first_slide_num =
                                String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            found = true;
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    if found {
        Ok(Some(info))
    } else {
        Ok(None)
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
                if e.name().as_ref() == b"a:xfrm" {
                    parse_transform(reader, &mut shape.transform, slide_path)?;
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"p:grpSpPr" {
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

    Ok(())
}

fn parse_transform(
    reader: &mut Reader<&[u8]>,
    transform: &mut ShapeTransform,
    slide_path: &str,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"a:off" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"x" => {
                                transform.x = String::from_utf8_lossy(&attr.value)
                                    .parse::<i64>()
                                    .unwrap_or(0)
                            }
                            b"y" => {
                                transform.y = String::from_utf8_lossy(&attr.value)
                                    .parse::<i64>()
                                    .unwrap_or(0)
                            }
                            _ => {}
                        }
                    }
                }
                b"a:ext" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"cx" => {
                                transform.width = String::from_utf8_lossy(&attr.value)
                                    .parse::<u64>()
                                    .unwrap_or(0)
                            }
                            b"cy" => {
                                transform.height = String::from_utf8_lossy(&attr.value)
                                    .parse::<u64>()
                                    .unwrap_or(0)
                            }
                            _ => {}
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"a:off" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"x" => {
                                transform.x = String::from_utf8_lossy(&attr.value)
                                    .parse::<i64>()
                                    .unwrap_or(0)
                            }
                            b"y" => {
                                transform.y = String::from_utf8_lossy(&attr.value)
                                    .parse::<i64>()
                                    .unwrap_or(0)
                            }
                            _ => {}
                        }
                    }
                }
                b"a:ext" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"cx" => {
                                transform.width = String::from_utf8_lossy(&attr.value)
                                    .parse::<u64>()
                                    .unwrap_or(0)
                            }
                            b"cy" => {
                                transform.height = String::from_utf8_lossy(&attr.value)
                                    .parse::<u64>()
                                    .unwrap_or(0)
                            }
                            _ => {}
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"a:xfrm" || e.name().as_ref() == b"p:xfrm" {
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

    Ok(())
}

fn parse_text_body(reader: &mut Reader<&[u8]>, slide_path: &str) -> Result<ShapeText, ParseError> {
    let mut paragraphs = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"a:p" {
                    let paragraph = parse_text_paragraph(reader, slide_path)?;
                    paragraphs.push(paragraph);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"p:txBody" {
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

    Ok(ShapeText { paragraphs })
}

fn parse_text_body_table(
    reader: &mut Reader<&[u8]>,
    slide_path: &str,
) -> Result<ShapeText, ParseError> {
    let mut paragraphs = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"a:p" {
                    let paragraph = parse_text_paragraph(reader, slide_path)?;
                    paragraphs.push(paragraph);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"a:txBody" {
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

    Ok(ShapeText { paragraphs })
}

fn shape_text_to_plain(text: &ShapeText) -> String {
    let mut out = String::new();
    for (p_idx, para) in text.paragraphs.iter().enumerate() {
        if p_idx > 0 {
            out.push('\n');
        }
        for run in &para.runs {
            out.push_str(&run.text);
        }
    }
    out
}

fn extract_c_sld_name(xml: &str) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref().ends_with(b"cSld") {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"name" {
                            return Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    None
}

fn parse_notes_slide(
    xml: &str,
    path: &str,
    rels: &Relationships,
    parser: &mut PptxParser,
    zip: &mut SecureZipReader<impl Read + Seek>,
) -> Result<(NotesSlide, String), ParseError> {
    let mut slide = NotesSlide::new();
    slide.span = Some(SourceSpan::new(path));

    let shapes = parser.parse_shapes_from_xml(xml, path, rels, zip)?;
    slide.shapes = shapes;

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Text(e)) => {
                let value = e.unescape().unwrap_or_default();
                if !text.is_empty() {
                    text.push(' ');
                }
                text.push_str(&value);
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    if !text.trim().is_empty() {
        slide.text = Some(text.clone());
    }

    Ok((slide, text))
}

fn parse_presentation_properties(
    xml: &str,
    path: &str,
) -> Result<PresentationProperties, ParseError> {
    let mut props = PresentationProperties::new();
    props.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"p:presentationPr" {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"autoCompressPictures" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                props.auto_compress_pictures =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"compatMode" => {
                                props.compat_mode =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            b"rtl" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                props.rtl = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"showSpecialPlsOnTitleSld" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                props.show_special_placeholders =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"removePersonalInfoOnSave" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                props.remove_personal_info_on_save =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"showInkAnnotation" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                props.show_ink_annotation =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(props)
}

fn parse_view_properties(xml: &str, path: &str) -> Result<ViewProperties, ParseError> {
    let mut props = ViewProperties::new();
    props.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"p:viewPr" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"lastView" => {
                                props.last_view =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            b"showComments" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                props.show_comments =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"showHiddenSlides" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                props.show_hidden_slides =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"showGuides" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                props.show_guides =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"showGrid" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                props.show_grid = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"showOutlineIcons" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                props.show_outline_icons =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            _ => {}
                        }
                    }
                }
                b"p:zoom" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"percent" {
                            props.zoom = String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(props)
}

fn parse_table_styles(xml: &str, path: &str) -> Result<TableStyleSet, ParseError> {
    let mut styles = TableStyleSet::new();
    styles.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"a:tblStyleLst" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"def" {
                            styles.default_style_id =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"a:tblStyle" => {
                    let mut style_id = None;
                    let mut name = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"styleId" => {
                                style_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            b"name" => {
                                name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            _ => {}
                        }
                    }
                    if let Some(style_id) = style_id {
                        styles.styles.push(TableStyle { style_id, name });
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(styles)
}

fn parse_comment_authors(xml: &str, path: &str) -> Result<Vec<PptxCommentAuthor>, ParseError> {
    let mut authors: Vec<PptxCommentAuthor> = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref().ends_with(b"cmAuthor") {
                    let mut author_id = None;
                    let mut name = None;
                    let mut initials = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"id" => {
                                author_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"name" => {
                                name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"initials" => {
                                initials = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            _ => {}
                        }
                    }
                    if let Some(author_id) = author_id {
                        authors.push(PptxCommentAuthor {
                            id: NodeId::new(),
                            author_id,
                            name,
                            initials,
                            span: Some(SourceSpan::new(path)),
                        });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(authors)
}

fn parse_comments(
    xml: &str,
    path: &str,
    authors: &HashMap<u32, (Option<String>, Option<String>)>,
) -> Result<Vec<PptxComment>, ParseError> {
    let mut comments: Vec<PptxComment> = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut current: Option<PptxComment> = None;
    let mut in_text = false;
    let mut text_buf = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref().ends_with(b"cm") {
                    let mut author_id = None;
                    let mut dt = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"authorId" => {
                                author_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"dt" => dt = Some(String::from_utf8_lossy(&attr.value).to_string()),
                            _ => {}
                        }
                    }
                    current = Some(PptxComment {
                        id: NodeId::new(),
                        author_id,
                        author_name: None,
                        author_initials: None,
                        datetime: dt,
                        text: String::new(),
                        span: Some(SourceSpan::new(path)),
                    });
                    text_buf.clear();
                } else if e.name().as_ref().ends_with(b"t") {
                    in_text = true;
                }
            }
            Ok(Event::Text(e)) => {
                if in_text {
                    text_buf.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref().ends_with(b"t") {
                    in_text = false;
                    if !text_buf.is_empty() {
                        if let Some(cur) = current.as_mut() {
                            if !cur.text.is_empty() {
                                cur.text.push(' ');
                            }
                            cur.text.push_str(&text_buf);
                        }
                        text_buf.clear();
                    }
                } else if e.name().as_ref().ends_with(b"cm") {
                    if let Some(mut cur) = current.take() {
                        if let Some(author_id) = cur.author_id {
                            if let Some((name, initials)) = authors.get(&author_id) {
                                cur.author_name = name.clone();
                                cur.author_initials = initials.clone();
                            }
                        }
                        comments.push(cur);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(comments)
}

fn parse_presentation_tags(xml: &str, path: &str) -> Result<Vec<PresentationTag>, ParseError> {
    let mut tags: Vec<PresentationTag> = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref().ends_with(b"tag") {
                    let mut name = None;
                    let mut val = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => {
                                name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"val" => val = Some(String::from_utf8_lossy(&attr.value).to_string()),
                            _ => {}
                        }
                    }
                    if let Some(name) = name {
                        tags.push(PresentationTag {
                            id: NodeId::new(),
                            name,
                            value: val,
                            span: Some(SourceSpan::new(path)),
                        });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(tags)
}

fn parse_smartart_part(xml: &str, path: &str) -> Result<SmartArtPart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut root = None;
    let mut point_count: u32 = 0;
    let mut connection_count: u32 = 0;
    let mut rel_ids: Vec<String> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if root.is_none() {
                    root = Some(String::from_utf8_lossy(e.name().as_ref()).to_string());
                }
                let name_buf = e.name().as_ref().to_vec();
                let name = name_buf.as_slice();
                if name.ends_with(b":pt") || name == b"dgm:pt" {
                    point_count += 1;
                }
                if name.ends_with(b":cxn") || name == b"dgm:cxn" {
                    connection_count += 1;
                }
                if name.ends_with(b":relIds") || name == b"dgm:relIds" {
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        if key == b"r:dm" || key == b"r:lo" || key == b"r:qs" || key == b"r:cs" {
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if !val.is_empty() {
                                rel_ids.push(val);
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    let kind = if path.contains("layout") {
        "layout"
    } else if path.contains("style") {
        "style"
    } else if path.contains("colors") {
        "colors"
    } else {
        "data"
    };

    Ok(SmartArtPart {
        id: NodeId::new(),
        kind: kind.to_string(),
        path: path.to_string(),
        root_element: root,
        point_count: if point_count > 0 {
            Some(point_count)
        } else {
            None
        },
        connection_count: if connection_count > 0 {
            Some(connection_count)
        } else {
            None
        },
        rel_ids,
        span: Some(SourceSpan::new(path)),
    })
}

fn parse_text_paragraph(
    reader: &mut Reader<&[u8]>,
    slide_path: &str,
) -> Result<ShapeTextParagraph, ParseError> {
    let mut runs = Vec::new();
    let mut alignment: Option<TextAlignment> = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"a:pPr" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"algn" {
                            alignment = map_alignment(&String::from_utf8_lossy(&attr.value));
                        }
                    }
                }
                b"a:r" => {
                    let run = parse_text_run(reader, slide_path)?;
                    runs.push(run);
                }
                b"a:br" => {
                    runs.push(ShapeTextRun {
                        text: "\n".to_string(),
                        bold: None,
                        italic: None,
                        font_size: None,
                        font_family: None,
                    });
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"a:p" {
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

    Ok(ShapeTextParagraph { runs, alignment })
}

fn parse_text_run(
    reader: &mut Reader<&[u8]>,
    slide_path: &str,
) -> Result<ShapeTextRun, ParseError> {
    let mut text = String::new();
    let mut bold = None;
    let mut italic = None;
    let mut font_size = None;
    let mut font_family = None;

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"a:t" => {
                    let t = reader.read_text(e.name()).map_err(|e| ParseError::Xml {
                        file: slide_path.to_string(),
                        message: e.to_string(),
                    })?;
                    text.push_str(&t);
                }
                b"a:rPr" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"b" => bold = Some(attr.value.as_ref() == b"1"),
                            b"i" => italic = Some(attr.value.as_ref() == b"1"),
                            b"sz" => {
                                font_size = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            _ => {}
                        }
                    }
                }
                b"a:latin" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"typeface" {
                            font_family = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"a:r" {
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

    Ok(ShapeTextRun {
        text,
        bold,
        italic,
        font_size,
        font_family,
    })
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
            b"spd" => transition.speed = Some(String::from_utf8_lossy(&attr.value).to_string()),
            b"advClick" => {
                let v = String::from_utf8_lossy(&attr.value);
                transition.advance_on_click = Some(v == "1" || v.eq_ignore_ascii_case("true"));
            }
            b"advTm" => {
                transition.advance_after_ms =
                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
            }
            b"dur" => {
                transition.duration_ms = String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
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

#[derive(Default)]
struct SlideMasterMeta {
    preserve: Option<bool>,
    show_master_sp: Option<bool>,
    show_master_ph_anim: Option<bool>,
}

fn parse_slide_master_meta(xml: &str, path: &str) -> Result<SlideMasterMeta, ParseError> {
    let mut meta = SlideMasterMeta::default();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"p:sldMaster" {
                    for attr in e.attributes().flatten() {
                        let v = String::from_utf8_lossy(&attr.value);
                        match attr.key.as_ref() {
                            b"preserve" => {
                                meta.preserve = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"showMasterSp" => {
                                meta.show_master_sp =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"showMasterPhAnim" => {
                                meta.show_master_ph_anim =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            _ => {}
                        }
                    }
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(meta)
}

#[derive(Default)]
struct SlideLayoutMeta {
    layout_type: Option<String>,
    matching_name: Option<String>,
    preserve: Option<bool>,
    show_master_sp: Option<bool>,
    show_master_ph_anim: Option<bool>,
}

fn parse_slide_layout_meta(xml: &str, path: &str) -> Result<SlideLayoutMeta, ParseError> {
    let mut meta = SlideLayoutMeta::default();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"p:sldLayout" {
                    for attr in e.attributes().flatten() {
                        let v = String::from_utf8_lossy(&attr.value);
                        match attr.key.as_ref() {
                            b"type" => meta.layout_type = Some(v.to_string()),
                            b"matchingName" => meta.matching_name = Some(v.to_string()),
                            b"preserve" => {
                                meta.preserve = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"showMasterSp" => {
                                meta.show_master_sp =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"showMasterPhAnim" => {
                                meta.show_master_ph_anim =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            _ => {}
                        }
                    }
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(meta)
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
                if matches!(
                    name.as_slice(),
                    b"p:anim"
                        | b"p:animEffect"
                        | b"p:animMotion"
                        | b"p:animRot"
                        | b"p:animScale"
                        | b"p:seq"
                ) {
                    let mut anim = SlideAnimation {
                        animation_type: String::from_utf8_lossy(&name).to_string(),
                        target: None,
                        duration_ms: None,
                        preset_id: None,
                        preset_class: None,
                        media_asset: None,
                    };
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"dur" => {
                                anim.duration_ms =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            }
                            b"presetID" => {
                                anim.preset_id =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            b"presetClass" => {
                                anim.preset_class =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            _ => {}
                        }
                    }
                    animations.push(anim);
                    current_index = Some(animations.len() - 1);
                } else if name.as_slice().ends_with(b":audio")
                    || name.as_slice().ends_with(b":video")
                {
                    let mut anim = SlideAnimation {
                        animation_type: String::from_utf8_lossy(&name).to_string(),
                        target: None,
                        duration_ms: None,
                        preset_id: None,
                        preset_class: None,
                        media_asset: None,
                    };
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"r:link" | b"r:embed" => {
                                let rel_id = String::from_utf8_lossy(&attr.value).to_string();
                                if let Some(rel) = relationships.get(&rel_id) {
                                    anim.target = Some(Relationships::resolve_target(
                                        slide_path,
                                        &rel.target,
                                    ));
                                } else {
                                    anim.target = Some(rel_id);
                                }
                            }
                            b"dur" => {
                                anim.duration_ms =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            }
                            _ => {}
                        }
                    }
                    animations.push(anim);
                    current_index = Some(animations.len() - 1);
                } else if name.as_slice() == b"p:spTgt" {
                    if let Some(idx) = current_index {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"spid" {
                                animations[idx].target =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"p:spTgt" {
                    if let Some(idx) = current_index {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"spid" {
                                animations[idx].target =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                } else if e.name().as_ref().ends_with(b":audio")
                    || e.name().as_ref().ends_with(b":video")
                {
                    let mut anim = SlideAnimation {
                        animation_type: String::from_utf8_lossy(e.name().as_ref()).to_string(),
                        target: None,
                        duration_ms: None,
                        preset_id: None,
                        preset_class: None,
                        media_asset: None,
                    };
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"r:link" | b"r:embed" => {
                                let rel_id = String::from_utf8_lossy(&attr.value).to_string();
                                if let Some(rel) = relationships.get(&rel_id) {
                                    anim.target = Some(Relationships::resolve_target(
                                        slide_path,
                                        &rel.target,
                                    ));
                                } else {
                                    anim.target = Some(rel_id);
                                }
                            }
                            b"dur" => {
                                anim.duration_ms =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            }
                            _ => {}
                        }
                    }
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

fn map_shape_type(value: &str) -> ShapeType {
    match value {
        "rect" => ShapeType::Rectangle,
        "roundRect" => ShapeType::RoundRect,
        "ellipse" => ShapeType::Ellipse,
        "triangle" => ShapeType::Triangle,
        "line" => ShapeType::Line,
        "arrow" => ShapeType::Arrow,
        _ => ShapeType::Custom,
    }
}

fn map_alignment(value: &str) -> Option<TextAlignment> {
    match value {
        "l" => Some(TextAlignment::Left),
        "ctr" => Some(TextAlignment::Center),
        "r" => Some(TextAlignment::Right),
        "just" => Some(TextAlignment::Justify),
        "dist" => Some(TextAlignment::Distribute),
        _ => None,
    }
}

fn classify_relationship(rel_type_uri: &str) -> ExternalRefType {
    if rel_type_uri.contains("hyperlink") {
        ExternalRefType::Hyperlink
    } else if rel_type_uri.contains("image") {
        ExternalRefType::Image
    } else if rel_type_uri.contains("slideMaster") || rel_type_uri.contains("slideLayout") {
        ExternalRefType::SlideMaster
    } else if rel_type_uri.contains("oleObject") {
        ExternalRefType::OleLink
    } else if rel_type_uri.contains("external") {
        ExternalRefType::DataConnection
    } else {
        ExternalRefType::Other
    }
}

fn get_rels_path(part_path: &str) -> String {
    if let Some(idx) = part_path.rfind('/') {
        let dir = &part_path[..idx + 1];
        let file = &part_path[idx + 1..];
        format!("{}_rels/{}.rels", dir, file)
    } else {
        format!("_rels/{}.rels", part_path)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::zip_handler::SecureZipReader;
    use docir_core::ir::IRNode;
    use std::io::Cursor;

    #[test]
    fn test_parse_slide_list() {
        let xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId r:id="rId1"/>
            <p:sldId r:id="rId2"/>
          </p:sldIdLst>
        </p:presentation>
        "#;

        let slides = parse_slide_list(xml).expect("parse slide list");
        assert_eq!(slides, vec!["rId1", "rId2"]);
    }

    #[test]
    fn test_parse_presentation_info() {
        let xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        firstSlideNum="5">
          <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
          <p:notesSz cx="6858000" cy="9144000"/>
          <p:showPr showType="kiosk" loop="1" showNarration="0" showAnimation="1" useTimings="1"/>
        </p:presentation>
        "#;
        let info = parse_presentation_info(xml, "ppt/presentation.xml")
            .expect("info")
            .expect("info present");
        assert_eq!(info.first_slide_num, Some(5));
        assert_eq!(info.slide_size.as_ref().unwrap().cx, 9144000);
        assert_eq!(info.notes_size.as_ref().unwrap().cy, 9144000);
        assert_eq!(info.show_type.as_deref(), Some("kiosk"));
        assert_eq!(info.show_loop, Some(true));
        assert_eq!(info.show_narration, Some(false));
        assert_eq!(info.show_animation, Some(true));
        assert_eq!(info.use_timings, Some(true));
    }

    #[test]
    fn test_parse_presentation_and_view_properties_extended() {
        let pres_xml = r#"
        <p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                          removePersonalInfoOnSave="1"
                          showInkAnnotation="0"/>
        "#;
        let view_xml = r#"
        <p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                  lastView="slideSorterView"
                  showHiddenSlides="1"
                  showGuides="0"
                  showGrid="1"
                  showOutlineIcons="1">
          <p:zoom percent="85"/>
        </p:viewPr>
        "#;
        let props =
            parse_presentation_properties(pres_xml, "ppt/presProps.xml").expect("pres props");
        assert_eq!(props.remove_personal_info_on_save, Some(true));
        assert_eq!(props.show_ink_annotation, Some(false));

        let view = parse_view_properties(view_xml, "ppt/viewProps.xml").expect("view props");
        assert_eq!(view.last_view.as_deref(), Some("slideSorterView"));
        assert_eq!(view.show_hidden_slides, Some(true));
        assert_eq!(view.show_guides, Some(false));
        assert_eq!(view.show_grid, Some(true));
        assert_eq!(view.show_outline_icons, Some(true));
        assert_eq!(view.zoom, Some(85));
    }

    #[test]
    fn test_parse_slide_shapes() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               show="0">
          <p:cSld name="Title Slide">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="1" name="Title"/>
                </p:nvSpPr>
                <p:spPr>
                  <a:xfrm>
                    <a:off x="100" y="200"/>
                    <a:ext cx="300" cy="400"/>
                  </a:xfrm>
                </p:spPr>
                <p:txBody>
                  <a:p>
                    <a:r>
                      <a:rPr b="1" sz="2400"/>
                      <a:t>Hello</a:t>
                    </a:r>
                  </a:p>
                </p:txBody>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;

        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let slide_id = parser
            .parse_slide(
                &mut zip,
                slide_xml,
                1,
                "ppt/slides/slide1.xml",
                &Relationships::default(),
                None,
                None,
            )
            .expect("parse slide");
        let store = parser.into_store();

        let slide = match store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };

        assert_eq!(slide.number, 1);
        assert!(slide.hidden);
        assert_eq!(slide.name.as_deref(), Some("Title Slide"));
        assert_eq!(slide.shapes.len(), 1);

        let shape = match store.get(slide.shapes[0]) {
            Some(IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };

        assert_eq!(shape.name.as_deref(), Some("Title"));
        assert_eq!(shape.transform.x, 100);
        assert_eq!(shape.transform.y, 200);
        assert_eq!(shape.transform.width, 300);
        assert_eq!(shape.transform.height, 400);
        assert!(shape.text.is_some());
    }

    #[test]
    fn test_parse_pic_with_media_target() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <p:cSld>
            <p:spTree>
              <p:pic>
                <p:nvPicPr>
                  <p:cNvPr id="2" name="Picture 1" descr="Alt text"/>
                </p:nvPicPr>
                <p:blipFill>
                  <a:blip r:embed="rId2"/>
                </p:blipFill>
                <p:spPr>
                  <a:xfrm>
                    <a:off x="10" y="20"/>
                    <a:ext cx="300" cy="400"/>
                  </a:xfrm>
                </p:spPr>
              </p:pic>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;

        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId2"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="../media/image2.png"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels parse");

        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let slide_id = parser
            .parse_slide(
                &mut zip,
                slide_xml,
                1,
                "ppt/slides/slide1.xml",
                &rels,
                None,
                None,
            )
            .expect("parse slide");
        let store = parser.into_store();

        let slide = match store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };
        let shape = match store.get(slide.shapes[0]) {
            Some(IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };
        assert_eq!(shape.shape_type, ShapeType::Picture);
        assert_eq!(shape.media_target.as_deref(), Some("ppt/media/image2.png"));
        assert_eq!(shape.alt_text.as_deref(), Some("Alt text"));
    }

    #[test]
    fn test_parse_graphic_frame_chart() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <p:cSld>
            <p:spTree>
              <p:graphicFrame>
                <p:nvGraphicFramePr>
                  <p:cNvPr id="3" name="Chart 1"/>
                </p:nvGraphicFramePr>
                <p:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="1000" cy="800"/>
                </p:xfrm>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                    <c:chart r:id="rId3"/>
                  </a:graphicData>
                </a:graphic>
              </p:graphicFrame>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;

        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId3"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            Target="../charts/chart1.xml"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels parse");

        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let slide_id = parser
            .parse_slide(
                &mut zip,
                slide_xml,
                1,
                "ppt/slides/slide1.xml",
                &rels,
                None,
                None,
            )
            .expect("parse slide");
        let store = parser.into_store();

        let slide = match store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };
        let shape = match store.get(slide.shapes[0]) {
            Some(IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };
        assert_eq!(shape.shape_type, ShapeType::Chart);
        assert_eq!(shape.media_target.as_deref(), Some("ppt/charts/chart1.xml"));
    }

    #[test]
    fn test_parse_notes_slide_text() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld>
            <p:spTree/>
          </p:cSld>
        </p:sld>
        "#;
        let notes_xml = r#"
        <p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:sp>
                <p:txBody>
                  <a:p>
                    <a:r><a:t>First note</a:t></a:r>
                  </a:p>
                  <a:p>
                    <a:r><a:t>Second note</a:t></a:r>
                  </a:p>
                </p:txBody>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:notes>
        "#;

        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let notes_text = parse_notes_slide(
            notes_xml,
            "ppt/notesSlides/notesSlide1.xml",
            &Relationships::default(),
            &mut parser,
            &mut zip,
        )
        .unwrap()
        .1;
        let slide_id = parser
            .parse_slide(
                &mut zip,
                slide_xml,
                1,
                "ppt/slides/slide1.xml",
                &Relationships::default(),
                Some(&notes_text),
                None,
            )
            .expect("parse slide");
        let store = parser.into_store();

        let slide = match store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };
        assert_eq!(slide.notes.as_deref(), Some("First note Second note"));
    }

    #[test]
    fn test_parse_master_and_layout_shapes() {
        let master_xml = r#"
        <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                     preserve="1" showMasterSp="1" showMasterPhAnim="0">
          <p:cSld name="Master 1">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="1" name="MasterShape"/>
                </p:nvSpPr>
                <p:spPr>
                  <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="100" cy="100"/>
                  </a:xfrm>
                </p:spPr>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sldMaster>
        "#;
        let layout_xml = r#"
        <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                     type="title" matchingName="Title" preserve="1" showMasterSp="1" showMasterPhAnim="0">
          <p:cSld name="Layout 1">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="2" name="LayoutShape"/>
                </p:nvSpPr>
                <p:spPr>
                  <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="200" cy="200"/>
                  </a:xfrm>
                </p:spPr>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sldLayout>
        "#;

        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let master_shapes = parser
            .parse_shapes_from_xml(
                master_xml,
                "ppt/slideMasters/slideMaster1.xml",
                &Relationships::default(),
                &mut zip,
            )
            .expect("parse master shapes");
        let layout_id = parser
            .parse_slide_layout(
                layout_xml,
                "ppt/slideLayouts/slideLayout1.xml",
                &Relationships::default(),
                &mut zip,
            )
            .expect("parse layout");
        let mut master = docir_core::ir::SlideMaster::new();
        master.name = extract_c_sld_name(master_xml);
        let meta = parse_slide_master_meta(master_xml, "ppt/slideMasters/slideMaster1.xml")
            .expect("master meta");
        master.preserve = meta.preserve;
        master.show_master_sp = meta.show_master_sp;
        master.show_master_ph_anim = meta.show_master_ph_anim;
        master.shapes = master_shapes;
        master.layouts = vec![layout_id];
        let master_id = master.id;
        parser.store.insert(IRNode::SlideMaster(master));

        let store = parser.into_store();
        let master_node = match store.get(master_id) {
            Some(IRNode::SlideMaster(m)) => m,
            _ => panic!("missing master"),
        };
        assert_eq!(master_node.name.as_deref(), Some("Master 1"));
        assert_eq!(master_node.preserve, Some(true));
        assert_eq!(master_node.show_master_sp, Some(true));
        assert_eq!(master_node.show_master_ph_anim, Some(false));
        assert_eq!(master_node.shapes.len(), 1);
        assert_eq!(master_node.layouts.len(), 1);

        let layout_node = match store.get(layout_id) {
            Some(IRNode::SlideLayout(l)) => l,
            _ => panic!("missing layout"),
        };
        assert_eq!(layout_node.layout_type.as_deref(), Some("title"));
        assert_eq!(layout_node.matching_name.as_deref(), Some("Title"));
        assert_eq!(layout_node.preserve, Some(true));
        assert_eq!(layout_node.show_master_sp, Some(true));
        assert_eq!(layout_node.show_master_ph_anim, Some(false));
    }

    #[test]
    fn test_parse_notes_and_handout_master_shapes() {
        let notes_master_xml = r#"
        <p:notesMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld name="NotesMaster 1">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="10" name="NotesShape"/>
                </p:nvSpPr>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:notesMaster>
        "#;
        let handout_master_xml = r#"
        <p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld name="HandoutMaster 1">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="11" name="HandoutShape"/>
                </p:nvSpPr>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:handoutMaster>
        "#;

        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let notes_shapes = parser
            .parse_shapes_from_xml(
                notes_master_xml,
                "ppt/notesMasters/notesMaster1.xml",
                &Relationships::default(),
                &mut zip,
            )
            .expect("parse notes master shapes");
        let handout_shapes = parser
            .parse_shapes_from_xml(
                handout_master_xml,
                "ppt/handoutMasters/handoutMaster1.xml",
                &Relationships::default(),
                &mut zip,
            )
            .expect("parse handout master shapes");

        let mut notes_master = docir_core::ir::NotesMaster::new();
        notes_master.name = extract_c_sld_name(notes_master_xml);
        notes_master.shapes = notes_shapes;
        let notes_id = notes_master.id;
        parser.store.insert(IRNode::NotesMaster(notes_master));

        let mut handout_master = docir_core::ir::HandoutMaster::new();
        handout_master.name = extract_c_sld_name(handout_master_xml);
        handout_master.shapes = handout_shapes;
        let handout_id = handout_master.id;
        parser.store.insert(IRNode::HandoutMaster(handout_master));

        let store = parser.into_store();
        let notes = match store.get(notes_id) {
            Some(IRNode::NotesMaster(m)) => m,
            _ => panic!("missing notes master"),
        };
        let handout = match store.get(handout_id) {
            Some(IRNode::HandoutMaster(m)) => m,
            _ => panic!("missing handout master"),
        };
        assert_eq!(notes.name.as_deref(), Some("NotesMaster 1"));
        assert_eq!(notes.shapes.len(), 1);
        assert_eq!(handout.name.as_deref(), Some("HandoutMaster 1"));
        assert_eq!(handout.shapes.len(), 1);
    }

    #[test]
    fn test_parse_slide_transition_and_animation() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
          <p:transition spd="fast" advClick="1" advTm="500">
            <p:fade/>
          </p:transition>
          <p:timing>
            <p:tnLst>
              <p:par>
                <p:anim dur="300" presetID="1" presetClass="entr">
                  <p:tgtEl><p:spTgt spid="4"/></p:tgtEl>
                </p:anim>
              </p:par>
            </p:tnLst>
          </p:timing>
        </p:sld>
        "#;
        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let slide_id = parser
            .parse_slide(
                &mut zip,
                slide_xml,
                1,
                "ppt/slides/slide1.xml",
                &Relationships::default(),
                None,
                None,
            )
            .expect("slide");
        let store = parser.into_store();
        let slide = match store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };
        assert!(slide.transition.is_some());
        assert_eq!(slide.animations.len(), 1);
        assert_eq!(slide.animations[0].target.as_deref(), Some("4"));
    }

    #[test]
    fn test_parse_slide_timing_media() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld><p:spTree/></p:cSld>
          <p:timing>
            <p:audio r:link="rIdAudio" dur="5000"/>
            <p:video r:embed="rIdVideo" dur="12000"/>
          </p:timing>
        </p:sld>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdAudio"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
            Target="../media/audio1.wav"/>
          <Relationship Id="rIdVideo"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
            Target="../media/video1.mp4"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels");
        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let slide_id = parser
            .parse_slide(
                &mut zip,
                slide_xml,
                1,
                "ppt/slides/slide1.xml",
                &rels,
                None,
                None,
            )
            .expect("slide");
        let store = parser.into_store();
        let slide = match store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };
        assert_eq!(slide.animations.len(), 2);
        assert_eq!(
            slide.animations[0].target.as_deref(),
            Some("ppt/media/audio1.wav")
        );
        assert_eq!(
            slide.animations[1].target.as_deref(),
            Some("ppt/media/video1.mp4")
        );
    }

    #[test]
    fn test_parse_slide_comments() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
        </p:sld>
        "#;
        let comments_xml = r#"
        <p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cm authorId="1" dt="2024-01-01T00:00:00Z">
            <p:txBody><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Note</a:t></a:r></a:p></p:txBody>
          </p:cm>
        </p:cmLst>
        "#;
        let authors_xml = r#"
        <p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cmAuthor id="1" name="Alice" initials="AL"/>
        </p:cmAuthorLst>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdC"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            Target="../comments/comment1.xml"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels");
        let mut zip = build_zip_with_entries(vec![("ppt/comments/comment1.xml", comments_xml)]);
        let mut parser = PptxParser::new();
        let authors =
            parse_comment_authors(authors_xml, "ppt/commentAuthors.xml").expect("authors");
        parser.set_comment_authors(&authors);
        let slide_id = parser
            .parse_slide(
                &mut zip,
                slide_xml,
                1,
                "ppt/slides/slide1.xml",
                &rels,
                None,
                None,
            )
            .expect("slide");
        let store = parser.into_store();
        let slide = match store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };
        assert_eq!(slide.comments.len(), 1);
        let comment = match store.get(slide.comments[0]) {
            Some(IRNode::PptxComment(c)) => c,
            _ => panic!("missing comment"),
        };
        assert_eq!(comment.text, "Note");
        assert_eq!(comment.author_name.as_deref(), Some("Alice"));
        assert_eq!(comment.author_initials.as_deref(), Some("AL"));
    }

    #[test]
    fn test_parse_smartart_part_counts() {
        let xml = r#"
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <dgm:ptLst>
            <dgm:pt/>
            <dgm:pt/>
          </dgm:ptLst>
          <dgm:cxnLst>
            <dgm:cxn/>
          </dgm:cxnLst>
          <dgm:relIds r:dm="rId1" r:lo="rId2"/>
        </dgm:dataModel>
        "#;
        let part = parse_smartart_part(xml, "ppt/diagrams/data1.xml").expect("smartart");
        assert_eq!(part.point_count, Some(2));
        assert_eq!(part.connection_count, Some(1));
        assert_eq!(part.rel_ids.len(), 2);
        assert!(part.rel_ids.contains(&"rId1".to_string()));
        assert!(part.rel_ids.contains(&"rId2".to_string()));
    }

    #[test]
    fn test_parse_graphic_frame_ole_object() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld>
            <p:spTree>
              <p:graphicFrame>
                <p:nvGraphicFramePr>
                  <p:cNvPr id="5" name="OLE 1"/>
                </p:nvGraphicFramePr>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
                    <p:oleObj r:id="rId5"/>
                  </a:graphicData>
                </a:graphic>
              </p:graphicFrame>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId5"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
            Target="../embeddings/oleObject1.bin"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels parse");
        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let slide_id = parser
            .parse_slide(
                &mut zip,
                slide_xml,
                1,
                "ppt/slides/slide1.xml",
                &rels,
                None,
                None,
            )
            .expect("parse slide");
        let store = parser.into_store();

        let slide = match store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };
        assert_eq!(slide.shapes.len(), 1);
        let shape = match store.get(slide.shapes[0]) {
            Some(IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };
        assert_eq!(shape.shape_type, ShapeType::OleObject);
        assert_eq!(
            shape.media_target.as_deref(),
            Some("ppt/embeddings/oleObject1.bin")
        );
    }

    #[test]
    fn test_parse_pic_external_media_reference() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld>
            <p:spTree>
              <p:pic>
                <p:blipFill>
                  <a:blip r:embed="rIdAudio"/>
                </p:blipFill>
              </p:pic>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdAudio"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
            Target="https://example.com/audio.wav"
            TargetMode="External"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels parse");
        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let slide_id = parser
            .parse_slide(
                &mut zip,
                slide_xml,
                1,
                "ppt/slides/slide1.xml",
                &rels,
                None,
                None,
            )
            .expect("parse slide");
        let store = parser.into_store();

        let slide = match store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };
        let shape = match store.get(slide.shapes[0]) {
            Some(IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };
        assert_eq!(shape.shape_type, ShapeType::Audio);
        assert_eq!(
            shape.media_target.as_deref(),
            Some("https://example.com/audio.wav")
        );
    }

    #[test]
    fn test_parse_pic_linked_media_reference() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld>
            <p:spTree>
              <p:pic>
                <p:blipFill>
                  <a:blip r:link="rIdVideo"/>
                </p:blipFill>
              </p:pic>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdVideo"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
            Target="https://example.com/video.mp4"
            TargetMode="External"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels parse");
        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let slide_id = parser
            .parse_slide(
                &mut zip,
                slide_xml,
                1,
                "ppt/slides/slide1.xml",
                &rels,
                None,
                None,
            )
            .expect("parse slide");
        let store = parser.into_store();

        let slide = match store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };
        let shape = match store.get(slide.shapes[0]) {
            Some(IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };
        assert_eq!(shape.shape_type, ShapeType::Video);
        assert_eq!(
            shape.media_target.as_deref(),
            Some("https://example.com/video.mp4")
        );
    }

    #[test]
    fn test_parse_pic_embed_and_link_external() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld>
            <p:spTree>
              <p:pic>
                <p:blipFill>
                  <a:blip r:embed="rIdImg" r:link="rIdExt"/>
                </p:blipFill>
              </p:pic>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdImg"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="../media/image2.png"/>
          <Relationship Id="rIdExt"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
            Target="https://example.com/video.mp4"
            TargetMode="External"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels parse");
        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let slide_id = parser
            .parse_slide(
                &mut zip,
                slide_xml,
                1,
                "ppt/slides/slide1.xml",
                &rels,
                None,
                None,
            )
            .expect("parse slide");
        let store = parser.into_store();

        let slide = match store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };
        let shape = match store.get(slide.shapes[0]) {
            Some(IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };
        assert_eq!(shape.shape_type, ShapeType::Picture);
        assert_eq!(shape.media_target.as_deref(), Some("ppt/media/image2.png"));

        let mut ext_targets = Vec::new();
        for node in store.values() {
            if let IRNode::ExternalReference(ext) = node {
                ext_targets.push(ext.target.clone());
            }
        }
        assert!(ext_targets
            .iter()
            .any(|t| t == "https://example.com/video.mp4"));
    }

    #[test]
    fn test_parse_table_in_graphic_frame() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:graphicFrame>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                    <a:tbl>
                      <a:tblGrid>
                        <a:gridCol w="3000"/>
                        <a:gridCol w="3000"/>
                      </a:tblGrid>
                      <a:tr>
                        <a:tc><a:txBody><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody></a:tc>
                        <a:tc><a:txBody><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody></a:tc>
                      </a:tr>
                    </a:tbl>
                  </a:graphicData>
                </a:graphic>
              </p:graphicFrame>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;
        let mut parser = PptxParser::new();
        let mut zip = build_empty_zip();
        let slide_id = parser
            .parse_slide(
                &mut zip,
                slide_xml,
                1,
                "ppt/slides/slide1.xml",
                &Relationships::default(),
                None,
                None,
            )
            .expect("slide");
        let store = parser.into_store();
        let slide = match store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };
        let shape = match store.get(slide.shapes[0]) {
            Some(IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };
        assert_eq!(shape.shape_type, ShapeType::Table);
        let table_id = shape.table.expect("table id");
        let table = match store.get(table_id) {
            Some(IRNode::Table(t)) => t,
            _ => panic!("missing table"),
        };
        assert_eq!(table.rows.len(), 1);
        assert_eq!(table.grid.len(), 2);
    }

    fn build_empty_zip() -> SecureZipReader<Cursor<Vec<u8>>> {
        build_zip_with_entries(Vec::new())
    }

    fn build_zip_with_entries(entries: Vec<(&str, &str)>) -> SecureZipReader<Cursor<Vec<u8>>> {
        let mut data = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(&mut data));
            let options = zip::write::FileOptions::<()>::default();
            for (path, contents) in entries {
                writer.start_file(path, options).expect("start file");
                use std::io::Write;
                writer.write_all(contents.as_bytes()).expect("write file");
            }
            writer.finish().expect("finish zip");
        }
        SecureZipReader::new(std::io::Cursor::new(data), Default::default()).expect("zip")
    }
}
