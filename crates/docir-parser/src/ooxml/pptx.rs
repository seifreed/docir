//! PPTX presentation and slide parsing.

use crate::diagnostics::push_warning;
use crate::error::ParseError;
use crate::ooxml::part_utils::{
    parse_xml_part_with_span, read_xml_part_and_rels, read_xml_part_and_rels_optional,
};
use crate::ooxml::relationships::{rel_type, Relationship, Relationships, TargetMode};
use crate::xml_utils::read_event;
use crate::zip_handler::PackageReader;
use docir_core::ir::{
    Diagnostics, Document, GridColumn, IRNode, NotesSlide, Paragraph, PptxCommentAuthor,
    PresentationInfo, PresentationProperties, PresentationTag, Run, Shape, ShapeTransform,
    ShapeType, Slide, SlideAnimation, SlideSize, SlideTransition, SmartArtPart, Table, TableCell,
    TableRow, TableStyle, TableStyleSet, ViewProperties,
};
use docir_core::security::{ExternalRefType, ExternalReference, SecurityInfo};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::{HashMap, HashSet};

mod builder;
mod comments;
mod graphic_frame;
mod metadata;
mod presentation_parts;
mod relationships;
mod shapes;
mod slide;
mod text;
mod transform;

use comments::{parse_comment_authors, parse_comments};
use metadata::{
    map_shape_type, parse_presentation_properties, parse_presentation_tags,
    parse_slide_layout_meta, parse_slide_master_meta, parse_smartart_part, parse_table_styles,
    parse_view_properties,
};
use relationships::classify_relationship;
use shapes::parse_shape_properties;
use text::{parse_text_body, parse_text_body_table, shape_text_to_plain};
use transform::parse_transform;

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

    /// Returns the IR store.
    pub fn into_store(self) -> IrStore {
        self.store
    }

    fn parse_shapes_from_xml(
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
                Event::Start(e) => match e.name().as_ref() {
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
                Event::Eof => break,
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
    zip: &mut impl PackageReader,
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

#[cfg(test)]
mod tests;
