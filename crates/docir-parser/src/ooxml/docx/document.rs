//! DOCX document parsing (minimal but real).

use super::field::parse_field_instruction;
use crate::error::ParseError;
use crate::ooxml::relationships::{Relationships, TargetMode};
use crate::ooxml::shared::normalize_docx_target;
use crate::xml_utils::{attr_value, read_event, reader_from_str};
use docir_core::ir::{
    Border, BorderStyle, CommentRangeEnd, CommentRangeStart, CommentReference, Document, Field,
    Footer, GlossaryDocument, Header, Hyperlink, LineSpacingRule, NumberingInfo, PageBorders,
    Paragraph, ParagraphProperties, Revision, RevisionType, Run, RunProperties,
    StyleParagraphProperties, StyleRunProperties, TextAlignment, UnderlineStyle,
    VerticalTextAlignment, WebSettings, WordSettings,
};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::HashMap;

mod body;
mod comments;
mod drawing;
mod font_table;
mod glossary;
mod numbering;
mod paragraph;
mod sections;
mod styles;
mod table;
use body::{parse_block_until, parse_body_sections};
use drawing::parse_drawing;
use glossary::parse_doc_part;
use paragraph::{parse_paragraph, parse_paragraph_simple};
use sections::{apply_section_refs, SectionRef};
use table::parse_table;

#[derive(Debug, Clone, Copy)]
pub enum NoteKind {
    Footnote,
    Endnote,
}

#[derive(Debug, Clone, Copy)]
pub enum HeaderFooterKind {
    Header,
    Footer,
}

/// DOCX parser with internal store.
pub struct DocxParser {
    store: IrStore,
}

struct ParagraphParse {
    id: NodeId,
    section_ref: Option<SectionRef>,
}

impl DocxParser {
    pub fn new() -> Self {
        Self {
            store: IrStore::new(),
        }
    }

    pub(crate) fn store_mut(&mut self) -> &mut IrStore {
        &mut self.store
    }

    pub fn into_store(self) -> IrStore {
        self.store
    }

    pub fn parse_document(
        &mut self,
        xml: &str,
        rels: &Relationships,
        header_footer_map: Option<&HashMap<String, NodeId>>,
    ) -> Result<NodeId, ParseError> {
        let mut doc = Document::new(DocumentFormat::WordProcessing);

        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();

        loop {
            let event = read_event(&mut reader, &mut buf, "word/document.xml")?;
            match event {
                Event::Start(e) => {
                    if e.name().as_ref() == b"w:body" {
                        let sections =
                            parse_body_sections(self, &mut reader, rels, header_footer_map)?;
                        for section in sections {
                            let section_id = section.id;
                            self.store.insert(docir_core::ir::IRNode::Section(section));
                            doc.content.push(section_id);
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        let doc_id = doc.id;
        self.store.insert(docir_core::ir::IRNode::Document(doc));
        Ok(doc_id)
    }

    pub fn parse_glossary_document(
        &mut self,
        xml: &str,
        rels: &Relationships,
    ) -> Result<NodeId, ParseError> {
        let mut glossary = GlossaryDocument::new();
        glossary.span = Some(SourceSpan::new("word/glossary/document.xml"));

        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();

        loop {
            let event = read_event(&mut reader, &mut buf, "word/glossary/document.xml")?;
            match event {
                Event::Start(e) => {
                    if e.name().as_ref() == b"w:docPart" {
                        let entry = parse_doc_part(self, &mut reader, rels)?;
                        let entry_id = entry.id;
                        self.store
                            .insert(docir_core::ir::IRNode::GlossaryEntry(entry));
                        glossary.entries.push(entry_id);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        let glossary_id = glossary.id;
        self.store
            .insert(docir_core::ir::IRNode::GlossaryDocument(glossary));
        Ok(glossary_id)
    }

    pub fn parse_header_footer(
        &mut self,
        xml: &str,
        path: &str,
        kind: HeaderFooterKind,
        rels: &Relationships,
    ) -> Result<NodeId, ParseError> {
        let mut reader = reader_from_str(xml);

        let end_tag = match kind {
            HeaderFooterKind::Header => b"w:hdr".as_ref(),
            HeaderFooterKind::Footer => b"w:ftr".as_ref(),
        };
        let content = parse_block_until(self, &mut reader, rels, end_tag)?;
        let node_id = match kind {
            HeaderFooterKind::Header => {
                let mut header = Header::new();
                header.content = content;
                header.span = Some(SourceSpan::new(path));
                let id = header.id;
                self.store.insert(docir_core::ir::IRNode::Header(header));
                id
            }
            HeaderFooterKind::Footer => {
                let mut footer = Footer::new();
                footer.content = content;
                footer.span = Some(SourceSpan::new(path));
                let id = footer.id;
                self.store.insert(docir_core::ir::IRNode::Footer(footer));
                id
            }
        };
        Ok(node_id)
    }

    pub fn parse_settings(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let settings = parse_settings_like(xml)?;
        let id = settings.id;
        self.store
            .insert(docir_core::ir::IRNode::WordSettings(settings));
        Ok(id)
    }

    pub fn parse_web_settings(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let settings = parse_settings_like(xml)?;
        let mut web = WebSettings::new();
        web.entries = settings.entries;
        let id = web.id;
        self.store.insert(docir_core::ir::IRNode::WebSettings(web));
        Ok(id)
    }
}

fn style_run_from_run_props(props: RunProperties) -> StyleRunProperties {
    StyleRunProperties {
        font_family: props.font_family,
        font_size: props.font_size,
        bold: props.bold,
        italic: props.italic,
        underline: props.underline,
        strike: props.strike,
        color: props.color,
        highlight: props.highlight,
        vertical_align: props.vertical_align,
        all_caps: props.all_caps,
        small_caps: props.small_caps,
    }
}

fn style_paragraph_from_paragraph_props(props: ParagraphProperties) -> StyleParagraphProperties {
    StyleParagraphProperties {
        alignment: props.alignment,
        indentation: props.indentation,
        spacing: props.spacing,
        outline_level: props.outline_level,
        numbering: props.numbering,
        borders: props.borders,
        keep_next: props.keep_next,
        keep_lines: props.keep_lines,
        page_break_before: props.page_break_before,
        widow_control: props.widow_control,
    }
}

#[cfg(test)]
mod tests;

struct RunParse {
    run_id: NodeId,
    text: String,
    has_instr: bool,
    field_char: Option<String>,
    embedded: Vec<NodeId>,
}

#[derive(Debug, Clone, Copy)]
enum SdtMode {
    Block,
    Inline,
}

fn parse_run(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<RunParse, ParseError> {
    let mut text = String::new();
    let mut props = RunProperties::default();
    let mut has_instr = false;
    let mut field_char = None;
    let mut embedded = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:rPr" => {
                    parse_run_properties(reader, &mut props)?;
                }
                b"w:drawing" => {
                    if let Some(shape_id) = parse_drawing(parser, reader, rels)? {
                        embedded.push(shape_id);
                    }
                }
                b"w:pict" => {
                    if let Some(shape_id) = parse_vml_pict(parser, reader, rels)? {
                        embedded.push(shape_id);
                    }
                }
                b"w:footnoteReference" => {
                    if let Some(id) = attr_value(&e, b"w:id") {
                        let field_id = insert_note_reference(
                            parser,
                            reader,
                            docir_core::ir::FieldKind::FootnoteRef,
                            id,
                        );
                        embedded.push(field_id);
                    }
                }
                b"w:endnoteReference" => {
                    if let Some(id) = attr_value(&e, b"w:id") {
                        let field_id = insert_note_reference(
                            parser,
                            reader,
                            docir_core::ir::FieldKind::EndnoteRef,
                            id,
                        );
                        embedded.push(field_id);
                    }
                }
                b"w:fldChar" => {
                    field_char = attr_value(&e, b"w:fldCharType");
                }
                b"w:t" | b"w:instrText" | b"w:delText" => {
                    let content = reader.read_text(e.name()).unwrap_or_default();
                    if e.name().as_ref() == b"w:instrText" {
                        has_instr = true;
                    }
                    text.push_str(&content);
                }
                b"w:tab" => text.push('\t'),
                _ => {}
            },
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"w:tab" {
                    text.push('\t');
                } else if e.name().as_ref() == b"w:fldChar" {
                    field_char = attr_value(&e, b"w:fldCharType");
                } else if e.name().as_ref() == b"w:footnoteReference" {
                    if let Some(id) = attr_value(&e, b"w:id") {
                        let field_id = insert_note_reference(
                            parser,
                            reader,
                            docir_core::ir::FieldKind::FootnoteRef,
                            id,
                        );
                        embedded.push(field_id);
                    }
                } else if e.name().as_ref() == b"w:endnoteReference" {
                    if let Some(id) = attr_value(&e, b"w:id") {
                        let field_id = insert_note_reference(
                            parser,
                            reader,
                            docir_core::ir::FieldKind::EndnoteRef,
                            id,
                        );
                        embedded.push(field_id);
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:r" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    let mut run = Run::new(text);
    run.properties = props;
    run.span = Some(span_from_reader(reader, "word/document.xml"));
    let run_text = run.text.clone();
    let id = run.id;
    parser.store.insert(docir_core::ir::IRNode::Run(run));
    Ok(RunParse {
        run_id: id,
        text: run_text,
        has_instr,
        field_char,
        embedded,
    })
}

fn parse_revision_inline(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart,
    change_type: RevisionType,
) -> Result<NodeId, ParseError> {
    let mut revision = Revision::new(change_type);
    revision.revision_id = attr_value(start, b"w:id");
    revision.author = attr_value(start, b"w:author");
    revision.date = attr_value(start, b"w:date");
    revision.span = Some(SourceSpan::new("word/document.xml"));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:r" => {
                    let run = parse_run(parser, reader, rels)?;
                    revision.content.push(run.run_id);
                    revision.content.extend(run.embedded);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if matches!(
                    e.name().as_ref(),
                    b"w:ins"
                        | b"w:del"
                        | b"w:moveFrom"
                        | b"w:moveTo"
                        | b"w:pPrChange"
                        | b"w:rPrChange"
                ) {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    let id = revision.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::Revision(revision));
    Ok(id)
}

fn parse_revision_block(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart,
    change_type: RevisionType,
) -> Result<NodeId, ParseError> {
    let mut revision = Revision::new(change_type);
    revision.revision_id = attr_value(start, b"w:id");
    revision.author = attr_value(start, b"w:author");
    revision.date = attr_value(start, b"w:date");
    revision.span = Some(SourceSpan::new("word/document.xml"));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:p" => {
                    let para_id = parse_paragraph_simple(parser, reader, rels)?;
                    revision.content.push(para_id);
                }
                b"w:tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    revision.content.push(table_id);
                }
                b"w:r" => {
                    let run = parse_run(parser, reader, rels)?;
                    revision.content.push(run.run_id);
                    revision.content.extend(run.embedded);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if matches!(
                    e.name().as_ref(),
                    b"w:ins"
                        | b"w:del"
                        | b"w:moveFrom"
                        | b"w:moveTo"
                        | b"w:pPrChange"
                        | b"w:rPrChange"
                ) {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    let id = revision.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::Revision(revision));
    Ok(id)
}

fn parse_sdt(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    mode: SdtMode,
) -> Result<NodeId, ParseError> {
    let mut control = docir_core::ir::ContentControl::new();
    control.span = Some(span_from_reader(reader, "word/document.xml"));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:sdtPr" => {
                    parse_sdt_properties(reader, &mut control)?;
                }
                b"w:sdtContent" => {
                    let content = match mode {
                        SdtMode::Block => parse_sdt_content_block(parser, reader, rels)?,
                        SdtMode::Inline => parse_sdt_content_inline(parser, reader, rels)?,
                    };
                    control.content.extend(content);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:sdt" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    let id = control.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::ContentControl(control));
    Ok(id)
}

fn parse_sdt_properties(
    reader: &mut Reader<&[u8]>,
    control: &mut docir_core::ir::ContentControl,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:tag" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        control.tag = Some(val);
                    }
                }
                b"w:alias" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        control.alias = Some(val);
                    }
                }
                b"w:id" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        control.sdt_id = Some(val);
                    }
                }
                b"w:comboBox" => control.control_type = Some("comboBox".to_string()),
                b"w:dropDownList" => control.control_type = Some("dropDownList".to_string()),
                b"w:date" => control.control_type = Some("date".to_string()),
                b"w:checkbox" => control.control_type = Some("checkbox".to_string()),
                b"w:text" => control.control_type = Some("text".to_string()),
                b"w:picture" => control.control_type = Some("picture".to_string()),
                b"w:dataBinding" => {
                    control.data_binding_xpath = attr_value(&e, b"w:xpath");
                    control.data_binding_store_item_id = attr_value(&e, b"w:storeItemID");
                    control.data_binding_prefix_mappings = attr_value(&e, b"w:prefixMappings");
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:sdtPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_sdt_content_block(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<Vec<NodeId>, ParseError> {
    let mut content = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:p" => {
                    let para_id = parse_paragraph_simple(parser, reader, rels)?;
                    content.push(para_id);
                }
                b"w:tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    content.push(table_id);
                }
                b"w:sdt" => {
                    let sdt_id = parse_sdt(parser, reader, rels, SdtMode::Block)?;
                    content.push(sdt_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:sdtContent" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(content)
}

fn parse_sdt_content_inline(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<Vec<NodeId>, ParseError> {
    let mut runs = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:r" => {
                    let run = parse_run(parser, reader, rels)?;
                    runs.push(run.run_id);
                    runs.extend(run.embedded);
                }
                b"w:hyperlink" => {
                    let link_id = parse_hyperlink(parser, reader, rels, &e)?;
                    runs.push(link_id);
                }
                b"w:fldSimple" => {
                    let instr = attr_value(&e, b"w:instr");
                    let field_id = parse_field(parser, reader, instr)?;
                    runs.push(field_id);
                }
                b"w:commentRangeStart" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentRangeStart::new(cid);
                        node.span = Some(SourceSpan::new("word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentRangeStart(node));
                        runs.push(node_id);
                    }
                }
                b"w:commentRangeEnd" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentRangeEnd::new(cid);
                        node.span = Some(SourceSpan::new("word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentRangeEnd(node));
                        runs.push(node_id);
                    }
                }
                b"w:commentReference" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentReference::new(cid);
                        node.span = Some(SourceSpan::new("word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentReference(node));
                        runs.push(node_id);
                    }
                }
                b"w:bookmarkStart" => {
                    if let Some(bm_id) = attr_value(&e, b"w:id") {
                        let mut bm = docir_core::ir::BookmarkStart::new(bm_id);
                        bm.name = attr_value(&e, b"w:name");
                        let bm_id = bm.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::BookmarkStart(bm));
                        runs.push(bm_id);
                    }
                }
                b"w:bookmarkEnd" => {
                    if let Some(bm_id) = attr_value(&e, b"w:id") {
                        let bm = docir_core::ir::BookmarkEnd::new(bm_id);
                        let bm_id = bm.id;
                        parser.store.insert(docir_core::ir::IRNode::BookmarkEnd(bm));
                        runs.push(bm_id);
                    }
                }
                b"w:ins" => {
                    let rev_id =
                        parse_revision_inline(parser, reader, rels, &e, RevisionType::Insert)?;
                    runs.push(rev_id);
                }
                b"w:del" => {
                    let rev_id =
                        parse_revision_inline(parser, reader, rels, &e, RevisionType::Delete)?;
                    runs.push(rev_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:sdtContent" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(runs)
}

fn parse_hyperlink(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart,
) -> Result<NodeId, ParseError> {
    let mut link = Hyperlink::new("", false);
    let mut rel_id_opt = None;
    if let Some(tooltip) = attr_value(start, b"w:tooltip") {
        link.tooltip = Some(tooltip);
    }
    if let Some(rel_id) = attr_value(start, b"r:id") {
        if let Some(rel) = rels.get(&rel_id) {
            link.target = rel.target.clone();
            link.is_external = rel.target_mode == TargetMode::External;
            link.relationship_id = Some(rel_id.clone());
            rel_id_opt = Some(rel_id);
        }
    }
    if let Some(anchor) = attr_value(start, b"w:anchor") {
        if link.target.is_empty() {
            link.target = format!("#{}", anchor);
        } else if !link.target.contains('#') {
            link.target = format!("{}#{}", link.target, anchor);
        }
    }
    let mut span = span_from_reader(reader, "word/document.xml");
    span.relationship_id = rel_id_opt.clone();
    link.span = Some(span);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"w:r" {
                    let run = parse_run(parser, reader, rels)?;
                    link.runs.push(run.run_id);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:hyperlink" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    let id = link.id;
    parser.store.insert(docir_core::ir::IRNode::Hyperlink(link));
    Ok(id)
}

fn parse_field(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    instruction: Option<String>,
) -> Result<NodeId, ParseError> {
    let mut field = Field::new(instruction);
    field.instruction_parsed = field
        .instruction
        .as_deref()
        .and_then(parse_field_instruction);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"w:r" {
                    let run = parse_run(parser, reader, &Relationships::default())?;
                    field.runs.push(run.run_id);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:fldSimple" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    let id = field.id;
    parser.store.insert(docir_core::ir::IRNode::Field(field));
    Ok(id)
}

fn parse_numbering(
    reader: &mut Reader<&[u8]>,
    props: &mut ParagraphProperties,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    let mut num_id = None;
    let mut level = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:numId" => {
                    num_id = attr_value(&e, b"w:val").and_then(|v| v.parse().ok());
                }
                b"w:ilvl" => {
                    level = attr_value(&e, b"w:val").and_then(|v| v.parse().ok());
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:numPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    if let (Some(num_id), Some(level)) = (num_id, level) {
        props.numbering = Some(NumberingInfo {
            num_id,
            level,
            format: None,
        });
    }
    Ok(())
}

fn parse_run_properties(
    reader: &mut Reader<&[u8]>,
    props: &mut RunProperties,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:rStyle" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.style_id = Some(val);
                    }
                }
                b"w:rFonts" => {
                    if let Some(val) = attr_value(&e, b"w:ascii") {
                        props.font_family = Some(val);
                    }
                }
                b"w:sz" => {
                    if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                        props.font_size = Some(val);
                    }
                }
                b"w:b" => props.bold = Some(true),
                b"w:i" => props.italic = Some(true),
                b"w:u" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.underline = match val.as_str() {
                            "double" => Some(UnderlineStyle::Double),
                            "thick" => Some(UnderlineStyle::Thick),
                            "dotted" => Some(UnderlineStyle::Dotted),
                            "dash" | "dashed" => Some(UnderlineStyle::Dashed),
                            _ => Some(UnderlineStyle::Single),
                        };
                    } else {
                        props.underline = Some(UnderlineStyle::Single);
                    }
                }
                b"w:strike" => props.strike = Some(true),
                b"w:color" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.color = Some(val);
                    }
                }
                b"w:highlight" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.highlight = Some(val);
                    }
                }
                b"w:caps" => {
                    props.all_caps = Some(bool_from_val(&e));
                }
                b"w:smallCaps" => {
                    props.small_caps = Some(bool_from_val(&e));
                }
                b"w:vertAlign" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.vertical_align = match val.as_str() {
                            "superscript" => Some(VerticalTextAlignment::Superscript),
                            "subscript" => Some(VerticalTextAlignment::Subscript),
                            _ => None,
                        };
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:rPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

pub(super) fn parse_border(start: &BytesStart) -> Option<Border> {
    let style = match attr_value(start, b"w:val").as_deref() {
        Some("nil") | Some("none") => BorderStyle::None,
        Some("single") => BorderStyle::Single,
        Some("double") => BorderStyle::Double,
        Some("thick") => BorderStyle::Thick,
        Some("dotted") => BorderStyle::Dotted,
        Some("dashed") => BorderStyle::Dashed,
        Some("dotDash") => BorderStyle::DotDash,
        Some("dotDotDash") => BorderStyle::DotDotDash,
        Some("triple") => BorderStyle::Triple,
        Some("wave") => BorderStyle::Wave,
        _ => BorderStyle::Single,
    };
    let width = attr_value(start, b"w:sz").and_then(|v| v.parse().ok());
    let color = attr_value(start, b"w:color").and_then(|v| {
        if v.eq_ignore_ascii_case("auto") {
            None
        } else {
            Some(v)
        }
    });
    Some(Border {
        style,
        width,
        color,
    })
}

fn insert_note_reference(
    parser: &mut DocxParser,
    reader: &Reader<&[u8]>,
    kind: docir_core::ir::FieldKind,
    note_id: String,
) -> NodeId {
    let mut field = Field::new(Some(match kind {
        docir_core::ir::FieldKind::FootnoteRef => "FOOTNOTE".to_string(),
        docir_core::ir::FieldKind::EndnoteRef => "ENDNOTE".to_string(),
        _ => "NOTE".to_string(),
    }));
    field.instruction_parsed = Some(docir_core::ir::FieldInstruction {
        kind,
        args: vec![note_id],
        switches: Vec::new(),
    });
    field.span = Some(span_from_reader(reader, "word/document.xml"));
    let id = field.id;
    parser.store.insert(docir_core::ir::IRNode::Field(field));
    id
}

fn parse_page_borders(reader: &mut Reader<&[u8]>) -> Result<Option<PageBorders>, ParseError> {
    let mut buf = Vec::new();
    let mut borders = PageBorders::default();
    let mut has_any = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let border = parse_border(&e);
                if border.is_none() {
                    continue;
                }
                match e.name().as_ref() {
                    b"w:top" => {
                        borders.top = border;
                        has_any = true;
                    }
                    b"w:bottom" => {
                        borders.bottom = border;
                        has_any = true;
                    }
                    b"w:left" => {
                        borders.left = border;
                        has_any = true;
                    }
                    b"w:right" => {
                        borders.right = border;
                        has_any = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:pgBorders" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    if has_any {
        Ok(Some(borders))
    } else {
        Ok(None)
    }
}

fn bool_from_val(start: &BytesStart) -> bool {
    match attr_value(start, b"w:val").as_deref() {
        Some("0") | Some("false") => false,
        _ => true,
    }
}

fn parse_vml_style_length(style: &str, key: &str) -> Option<i64> {
    parse_vml_style_length_value(style, key).map(|val| val.round() as i64)
}

fn parse_vml_style_length_u64(style: &str, key: &str) -> Option<u64> {
    parse_vml_style_length_value(style, key).and_then(|val| {
        if val >= 0.0 {
            Some(val.round() as u64)
        } else {
            None
        }
    })
}

fn parse_vml_style_length_value(style: &str, key: &str) -> Option<f64> {
    for part in style.split(';') {
        let mut iter = part.splitn(2, ':');
        let k = iter.next()?.trim();
        let v = iter.next()?.trim();
        if k.eq_ignore_ascii_case(key) {
            return parse_vml_length(v);
        }
    }
    None
}

fn parse_vml_length(value: &str) -> Option<f64> {
    let v = value.trim();
    if v.is_empty() {
        return None;
    }
    let mut split_idx = v.len();
    for (idx, ch) in v.char_indices().rev() {
        if ch.is_ascii_alphabetic() {
            split_idx = idx;
        } else {
            break;
        }
    }
    let (num_part, unit) = if split_idx < v.len() {
        v.split_at(split_idx)
    } else {
        (v, "")
    };
    let value = num_part.trim().parse::<f64>().ok()?;
    let unit = unit.trim();
    let emus = match unit {
        "" => value,
        "pt" => value * 12700.0,
        "in" => value * 914400.0,
        "cm" => value * 360000.0,
        "mm" => value * 36000.0,
        _ => return None,
    };
    Some(emus)
}

fn parse_vml_pict(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<Option<NodeId>, ParseError> {
    let mut buf = Vec::new();
    let mut rel_id: Option<String> = None;
    let mut name: Option<String> = None;
    let mut alt_text: Option<String> = None;
    let mut transform = docir_core::ir::ShapeTransform::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"v:imagedata" {
                    rel_id = attr_value(&e, b"r:id");
                } else if e.name().as_ref() == b"v:shape" {
                    name = attr_value(&e, b"name").or_else(|| attr_value(&e, b"id"));
                    alt_text = attr_value(&e, b"o:title").or_else(|| attr_value(&e, b"alt"));
                    if let Some(style) = attr_value(&e, b"style") {
                        if let Some(val) = parse_vml_style_length(&style, "left") {
                            transform.x = val;
                        }
                        if let Some(val) = parse_vml_style_length(&style, "top") {
                            transform.y = val;
                        }
                        if let Some(val) = parse_vml_style_length_u64(&style, "width") {
                            transform.width = val;
                        }
                        if let Some(val) = parse_vml_style_length_u64(&style, "height") {
                            transform.height = val;
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:pict" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    if let Some(rel_id) = rel_id {
        if let Some(rel) = rels.get(&rel_id) {
            let mut shape = docir_core::ir::Shape::new(docir_core::ir::ShapeType::Picture);
            shape.name = name;
            shape.alt_text = alt_text;
            shape.transform = transform;
            shape.relationship_id = Some(rel_id.clone());
            shape.media_target = Some(normalize_docx_target(&rel.target));
            let mut span = span_from_reader(reader, "word/document.xml");
            span.relationship_id = Some(rel_id.clone());
            shape.span = Some(span);
            let shape_id = shape.id;
            parser.store.insert(docir_core::ir::IRNode::Shape(shape));
            return Ok(Some(shape_id));
        }
    }
    Ok(None)
}

fn parse_settings_like(xml: &str) -> Result<WordSettings, ParseError> {
    let mut settings = WordSettings::new();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                let mut entry = docir_core::ir::SettingEntry {
                    name,
                    value: None,
                    attributes: Vec::new(),
                };
                for attr in e.attributes().flatten() {
                    let attr_name = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                    let attr_val = String::from_utf8_lossy(&attr.value).to_string();
                    entry.attributes.push(docir_core::ir::SettingAttribute {
                        name: attr_name,
                        value: attr_val,
                    });
                }
                settings.entries.push(entry);
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/settings.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(settings)
}

fn parse_num_abstract_id(reader: &mut Reader<&[u8]>) -> Result<u32, ParseError> {
    let mut buf = Vec::new();
    let mut abstract_id = 0u32;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"w:abstractNumId" {
                    if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                        abstract_id = val;
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:num" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/numbering.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(abstract_id)
}

fn xml_error(file: &str, err: quick_xml::Error) -> ParseError {
    ParseError::Xml {
        file: file.to_string(),
        message: err.to_string(),
    }
}

fn line_col(data: &[u8], pos: usize) -> Option<(u32, u32)> {
    if pos > data.len() {
        return None;
    }
    let slice = &data[..pos];
    let mut line = 1u32;
    let mut col = 1u32;
    for &b in slice {
        if b == b'\n' {
            line += 1;
            col = 1;
        } else {
            col = col.saturating_add(1);
        }
    }
    Some((line, col))
}

pub(super) fn span_from_reader(reader: &Reader<&[u8]>, file_path: &str) -> SourceSpan {
    let mut span = SourceSpan::new(file_path);
    if let Ok(pos) = usize::try_from(reader.buffer_position()) {
        if let Some((line, col)) = line_col(reader.get_ref(), pos) {
            span.line = Some(line);
            span.column = Some(col);
        }
    }
    span
}
