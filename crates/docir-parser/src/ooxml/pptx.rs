//! PPTX presentation and slide parsing.

use crate::diagnostics::push_warning;
use crate::error::ParseError;
use crate::ooxml::part_utils::{
    parse_xml_part_with_span, read_xml_part_and_rels, read_xml_part_and_rels_optional,
};
use crate::ooxml::relationships::{rel_type, Relationship, Relationships, TargetMode};
use crate::xml_utils::{attr_u32, attr_u64_from_bytes, attr_value, read_event, xml_error};
use crate::zip_handler::PackageReader;
use docir_core::ir::{
    Diagnostics, Document, GridColumn, IRNode, NotesSlide, Paragraph, PptxCommentAuthor,
    PresentationInfo, PresentationProperties, PresentationTag, Run, Shape, ShapeTransform,
    ShapeType, SlideSize, SmartArtPart, Table, TableCell, TableRow, TableStyle, TableStyleSet,
    ViewProperties,
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
mod graphic_frame_support;
mod metadata;
mod presentation_parts;
mod relationships;
#[path = "shape_parse.rs"]
mod shape_parse;
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
                        parse_grid_column(&e, &mut table);
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
                        parse_grid_column(&e, &mut table);
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"a:tbl" {
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
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"a:tc" {
                        let cell = self.parse_pptx_table_cell(reader, slide_path)?;
                        let id = cell.id;
                        self.store.insert(IRNode::TableCell(cell));
                        row.cells.push(id);
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"a:tr" {
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
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"a:txBody" {
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
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"a:tc" {
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

        Ok(cell)
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
                    if let Some(rel_id) = attr_value(&e, b"r:id") {
                        slide_ids.push(rel_id);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("ppt/presentation.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(slide_ids)
}

fn parse_grid_column(start: &BytesStart<'_>, table: &mut Table) {
    if let Some(width) = parse_u32_attr(start, b"w") {
        table.grid.push(GridColumn { width });
    }
}

fn parse_u64_attr(start: &BytesStart<'_>, key_name: &[u8]) -> Option<u64> {
    attr_u64_from_bytes(start, key_name)
}

fn parse_u32_attr(start: &BytesStart<'_>, key_name: &[u8]) -> Option<u32> {
    attr_u32(start, key_name)
}

fn parse_bool_attr(start: &BytesStart<'_>, key_name: &[u8]) -> Option<bool> {
    attr_value(start, key_name).map(|value| value == "1" || value.eq_ignore_ascii_case("true"))
}

fn parse_size_type_attr(start: &BytesStart<'_>) -> Option<String> {
    attr_value(start, b"type")
}

fn parse_size_attrs(start: &BytesStart<'_>) -> Option<(u64, u64)> {
    let cx = parse_u64_attr(start, b"cx");
    let cy = parse_u64_attr(start, b"cy");
    match (cx, cy) {
        (Some(cx), Some(cy)) => Some((cx, cy)),
        _ => None,
    }
}

fn apply_show_properties(start: &BytesStart<'_>, info: &mut PresentationInfo) {
    if let Some(value) = attr_value(start, b"showType") {
        info.show_type = Some(value);
    }
    if let Some(show_loop) = parse_bool_attr(start, b"loop") {
        info.show_loop = Some(show_loop);
    }
    if let Some(show_narration) = parse_bool_attr(start, b"showNarration") {
        info.show_narration = Some(show_narration);
    }
    if let Some(show_animation) = parse_bool_attr(start, b"showAnimation") {
        info.show_animation = Some(show_animation);
    }
    if let Some(use_timings) = parse_bool_attr(start, b"useTimings") {
        info.use_timings = Some(use_timings);
    }
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
                    if let Some((cx, cy)) = parse_size_attrs(&e) {
                        let size_type = parse_size_type_attr(&e);
                        info.slide_size = Some(SlideSize { cx, cy, size_type });
                        found = true;
                    }
                } else if name == b"p:notesSz" {
                    if let Some((cx, cy)) = parse_size_attrs(&e) {
                        info.notes_size = Some(SlideSize {
                            cx,
                            cy,
                            size_type: None,
                        });
                        found = true;
                    }
                } else if name == b"p:showPr" {
                    apply_show_properties(&e, &mut info);
                    found = true;
                } else if name == b"p:presentation" {
                    if let Some(first_slide_num) = parse_u32_attr(&e, b"firstSlideNum") {
                        info.first_slide_num = Some(first_slide_num);
                        found = true;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
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

fn extract_c_sld_name(xml: &str) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref().ends_with(b"cSld") {
                    return attr_value(&e, b"name");
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
                return Err(xml_error(path, e));
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
