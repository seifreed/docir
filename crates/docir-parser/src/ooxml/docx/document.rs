//! DOCX document parsing (minimal but real).

use crate::error::ParseError;
use crate::ooxml::relationships::{Relationships, TargetMode};
use docir_core::ir::{
    Border, BorderStyle, CellMargins, CellVerticalAlignment, Comment, CommentExtension,
    CommentExtensionSet, CommentIdMap, CommentIdMapEntry, CommentRangeEnd, CommentRangeStart,
    CommentReference, Document, Endnote, Field, FontEntry, FontTable, Footer, Footnote,
    GlossaryDocument, GlossaryEntry, Header, Hyperlink, LineNumberRestart, LineSpacingRule,
    MergeType, NumberingInfo, NumberingLevel, NumberingSet, PageBorders, PageMargins,
    PageOrientation, Paragraph, ParagraphProperties, Revision, RevisionType, RowHeight,
    RowHeightRule, Run, RunProperties, Section, SectionProperties, SectionType, Style,
    StyleParagraphProperties, StyleRunProperties, StyleSet, StyleType, Table, TableAlignment,
    TableBorders, TableCell, TableCellProperties, TableRow, TableWidth, TableWidthType,
    TextAlignment, UnderlineStyle, VerticalTextAlignment, WebSettings, WordSettings,
};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::HashMap;

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

struct BodyParse {
    content: Vec<NodeId>,
    headers: Vec<NodeId>,
    footers: Vec<NodeId>,
}

struct SectionRef {
    headers: Vec<NodeId>,
    footers: Vec<NodeId>,
    properties: SectionProperties,
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

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"w:body" => {
                        let sections =
                            parse_body_sections(self, &mut reader, rels, header_footer_map)?;
                        for section in sections {
                            let section_id = section.id;
                            self.store.insert(docir_core::ir::IRNode::Section(section));
                            doc.content.push(section_id);
                        }
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: "word/document.xml".to_string(),
                        message: e.to_string(),
                    });
                }
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

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"w:docPart" => {
                        let entry = parse_doc_part(self, &mut reader, rels)?;
                        let entry_id = entry.id;
                        self.store
                            .insert(docir_core::ir::IRNode::GlossaryEntry(entry));
                        glossary.entries.push(entry_id);
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: "word/glossary/document.xml".to_string(),
                        message: e.to_string(),
                    });
                }
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
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

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

    pub fn parse_styles(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let mut styles = StyleSet::new();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut current: Option<Style> = None;
        let mut in_name = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"w:style" => {
                        let style_id = attr_value(&e, b"w:styleId").unwrap_or_default();
                        let mut style = Style {
                            style_id,
                            name: None,
                            style_type: StyleType::Other,
                            based_on: None,
                            next: None,
                            is_default: attr_value(&e, b"w:default")
                                .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
                                .unwrap_or(false),
                            run_props: None,
                            paragraph_props: None,
                            table_props: None,
                        };
                        if let Some(t) = attr_value(&e, b"w:type") {
                            style.style_type = match t.as_str() {
                                "paragraph" => StyleType::Paragraph,
                                "character" => StyleType::Character,
                                "table" => StyleType::Table,
                                "numbering" => StyleType::Numbering,
                                _ => StyleType::Other,
                            };
                        }
                        current = Some(style);
                    }
                    b"w:name" => {
                        in_name = true;
                    }
                    b"w:rPr" => {
                        let mut props = RunProperties::default();
                        parse_run_properties(&mut reader, &mut props)?;
                        if let Some(style) = current.as_mut() {
                            style.run_props = Some(style_run_from_run_props(props));
                        }
                    }
                    b"w:pPr" => {
                        let mut para = Paragraph::new();
                        let _ = parse_paragraph_properties(&mut reader, &mut para, None)?;
                        if let Some(style) = current.as_mut() {
                            style.paragraph_props =
                                Some(style_paragraph_from_paragraph_props(para.properties));
                        }
                    }
                    b"w:tblPr" => {
                        if let Some(style) = current.as_mut() {
                            let mut props = docir_core::ir::TableProperties::default();
                            parse_table_properties(&mut reader, &mut props)?;
                            style.table_props = Some(props);
                        }
                    }
                    b"w:basedOn" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(style) = current.as_mut() {
                                style.based_on = Some(val);
                            }
                        }
                    }
                    b"w:next" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(style) = current.as_mut() {
                                style.next = Some(val);
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => match e.name().as_ref() {
                    b"w:name" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(style) = current.as_mut() {
                                style.name = Some(val);
                            }
                        }
                    }
                    b"w:basedOn" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(style) = current.as_mut() {
                                style.based_on = Some(val);
                            }
                        }
                    }
                    b"w:next" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(style) = current.as_mut() {
                                style.next = Some(val);
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::Text(e)) => {
                    if in_name {
                        if let Some(style) = current.as_mut() {
                            style.name = Some(e.unescape().unwrap_or_default().to_string());
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"w:name" {
                        in_name = false;
                    } else if e.name().as_ref() == b"w:style" {
                        if let Some(style) = current.take() {
                            styles.styles.push(style);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: "word/styles.xml".to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        let id = styles.id;
        self.store.insert(docir_core::ir::IRNode::StyleSet(styles));
        Ok(id)
    }

    pub fn parse_styles_with_effects(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let id = self.parse_styles(xml)?;
        if let Some(docir_core::ir::IRNode::StyleSet(set)) = self.store.get_mut(id) {
            set.with_effects = true;
        }
        Ok(id)
    }

    pub fn parse_numbering(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let mut set = NumberingSet::new();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut current_abs: Option<u32> = None;
        let mut current_levels: Vec<NumberingLevel> = Vec::new();
        let mut current_level: Option<NumberingLevel> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"w:abstractNum" => {
                        current_abs =
                            attr_value(&e, b"w:abstractNumId").and_then(|v| v.parse().ok());
                        current_levels.clear();
                    }
                    b"w:lvl" => {
                        let lvl = attr_value(&e, b"w:ilvl")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                        current_level = Some(NumberingLevel {
                            level: lvl,
                            format: None,
                            text: None,
                            start: None,
                            alignment: None,
                            suffix: None,
                            paragraph_props: None,
                            run_props: None,
                        });
                    }
                    b"w:numFmt" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.format = Some(val);
                            }
                        }
                    }
                    b"w:lvlText" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.text = Some(val);
                            }
                        }
                    }
                    b"w:start" => {
                        if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                            if let Some(level) = current_level.as_mut() {
                                level.start = Some(val);
                            }
                        }
                    }
                    b"w:lvlJc" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.alignment = match val.as_str() {
                                    "center" => Some(TextAlignment::Center),
                                    "right" => Some(TextAlignment::Right),
                                    "justify" => Some(TextAlignment::Justify),
                                    "distribute" => Some(TextAlignment::Distribute),
                                    _ => Some(TextAlignment::Left),
                                };
                            }
                        }
                    }
                    b"w:suff" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.suffix = Some(val);
                            }
                        }
                    }
                    b"w:pPr" => {
                        let mut para = Paragraph::new();
                        let _ = parse_paragraph_properties(&mut reader, &mut para, None)?;
                        if let Some(level) = current_level.as_mut() {
                            level.paragraph_props =
                                Some(style_paragraph_from_paragraph_props(para.properties));
                        }
                    }
                    b"w:rPr" => {
                        let mut props = RunProperties::default();
                        parse_run_properties(&mut reader, &mut props)?;
                        if let Some(level) = current_level.as_mut() {
                            level.run_props = Some(style_run_from_run_props(props));
                        }
                    }
                    b"w:num" => {
                        let num_id = attr_value(&e, b"w:numId")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                        let abstract_id = parse_num_abstract_id(&mut reader)?;
                        set.nums.push(docir_core::ir::NumInstance {
                            num_id,
                            abstract_id,
                        });
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => match e.name().as_ref() {
                    b"w:numFmt" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.format = Some(val);
                            }
                        }
                    }
                    b"w:lvlText" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.text = Some(val);
                            }
                        }
                    }
                    b"w:start" => {
                        if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                            if let Some(level) = current_level.as_mut() {
                                level.start = Some(val);
                            }
                        }
                    }
                    b"w:lvlJc" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.alignment = match val.as_str() {
                                    "center" => Some(TextAlignment::Center),
                                    "right" => Some(TextAlignment::Right),
                                    "justify" => Some(TextAlignment::Justify),
                                    "distribute" => Some(TextAlignment::Distribute),
                                    _ => Some(TextAlignment::Left),
                                };
                            }
                        }
                    }
                    b"w:suff" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(level) = current_level.as_mut() {
                                level.suffix = Some(val);
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => match e.name().as_ref() {
                    b"w:lvl" => {
                        if let Some(level) = current_level.take() {
                            current_levels.push(level);
                        }
                    }
                    b"w:abstractNum" => {
                        if let Some(abs_id) = current_abs.take() {
                            set.abstract_nums.push(docir_core::ir::AbstractNum {
                                abstract_id: abs_id,
                                levels: current_levels.clone(),
                            });
                        }
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: "word/numbering.xml".to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        let id = set.id;
        self.store.insert(docir_core::ir::IRNode::NumberingSet(set));
        Ok(id)
    }

    pub fn parse_comments(
        &mut self,
        xml: &str,
        rels: &Relationships,
    ) -> Result<Vec<NodeId>, ParseError> {
        parse_comments_like(self, xml, rels, None)
    }

    pub fn parse_notes(
        &mut self,
        xml: &str,
        kind: NoteKind,
        rels: &Relationships,
    ) -> Result<Vec<NodeId>, ParseError> {
        parse_comments_like(self, xml, rels, Some(kind))
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

    pub fn parse_font_table(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let mut table = FontTable::new();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut current: Option<FontEntry> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"w:font" {
                        let name = attr_value(&e, b"w:name").unwrap_or_default();
                        current = Some(FontEntry {
                            name,
                            alt_name: None,
                            charset: None,
                            family: None,
                            panose: None,
                        });
                    }
                }
                Ok(Event::Empty(e)) => match e.name().as_ref() {
                    b"w:altName" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(font) = current.as_mut() {
                                font.alt_name = Some(val);
                            }
                        }
                    }
                    b"w:charset" => {
                        if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                            if let Some(font) = current.as_mut() {
                                font.charset = Some(val);
                            }
                        }
                    }
                    b"w:family" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(font) = current.as_mut() {
                                font.family = Some(val);
                            }
                        }
                    }
                    b"w:panose1" => {
                        if let Some(val) = attr_value(&e, b"w:val") {
                            if let Some(font) = current.as_mut() {
                                font.panose = Some(val);
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"w:font" {
                        if let Some(font) = current.take() {
                            table.fonts.push(font);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: "word/fontTable.xml".to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        let id = table.id;
        self.store.insert(docir_core::ir::IRNode::FontTable(table));
        Ok(id)
    }

    pub fn parse_comments_extended(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let mut set = CommentExtensionSet::new();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"w16cid:commentExt" {
                        let comment_id = attr_value(&e, b"w:id").unwrap_or_default();
                        let entry = CommentExtension {
                            comment_id,
                            para_id: attr_value(&e, b"w16cid:paraId"),
                            parent_para_id: attr_value(&e, b"w16cid:parentParaId"),
                            done: attr_value(&e, b"w:done")
                                .map(|v| v == "1" || v.eq_ignore_ascii_case("true")),
                        };
                        set.entries.push(entry);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: "word/commentsExtended.xml".to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        let id = set.id;
        self.store
            .insert(docir_core::ir::IRNode::CommentExtensionSet(set));
        Ok(id)
    }

    pub fn parse_comments_ids(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let mut map = CommentIdMap::new();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"w16cid:commentId" {
                        let entry = CommentIdMapEntry {
                            comment_id: attr_value(&e, b"w:id").unwrap_or_default(),
                            para_id: attr_value(&e, b"w16cid:paraId"),
                            parent_para_id: attr_value(&e, b"w16cid:parentParaId"),
                        };
                        map.mappings.push(entry);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: "word/commentsIds.xml".to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        let id = map.id;
        self.store.insert(docir_core::ir::IRNode::CommentIdMap(map));
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

fn parse_body(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<BodyParse, ParseError> {
    let mut content = Vec::new();
    let mut headers = Vec::new();
    let mut footers = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:p" => {
                    let para = parse_paragraph(parser, reader, rels, header_footer_map)?;
                    content.push(para.id);
                }
                b"w:tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    content.push(table_id);
                }
                b"w:sectPr" => {
                    let section_ref = apply_section_refs(reader, header_footer_map)?;
                    headers.extend(section_ref.headers);
                    footers.extend(section_ref.footers);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:body"
                    || e.name().as_ref() == b"w:hdr"
                    || e.name().as_ref() == b"w:ftr"
                {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(BodyParse {
        content,
        headers,
        footers,
    })
}

fn parse_body_sections(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<Vec<Section>, ParseError> {
    let mut sections: Vec<Section> = Vec::new();
    let mut current = Section::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:p" => {
                    let para = parse_paragraph(parser, reader, rels, header_footer_map)?;
                    current.content.push(para.id);
                    if let Some(section_ref) = para.section_ref {
                        current.headers = section_ref.headers;
                        current.footers = section_ref.footers;
                        current.properties = section_ref.properties;
                        sections.push(current);
                        current = Section::new();
                    }
                }
                b"w:tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    current.content.push(table_id);
                }
                b"w:sdt" => {
                    let sdt_id = parse_sdt(parser, reader, rels, SdtMode::Block)?;
                    current.content.push(sdt_id);
                }
                b"w:sectPr" => {
                    let section_ref = apply_section_refs(reader, header_footer_map)?;
                    current.headers = section_ref.headers;
                    current.footers = section_ref.footers;
                    current.properties = section_ref.properties;
                    sections.push(current);
                    current = Section::new();
                }
                b"w:ins" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::Insert)?;
                    current.content.push(rev_id);
                }
                b"w:del" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::Delete)?;
                    current.content.push(rev_id);
                }
                b"w:moveFrom" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::MoveFrom)?;
                    current.content.push(rev_id);
                }
                b"w:moveTo" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::MoveTo)?;
                    current.content.push(rev_id);
                }
                b"w:pPrChange" | b"w:rPrChange" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::FormatChange)?;
                    current.content.push(rev_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:body" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    if !current.content.is_empty()
        || !current.headers.is_empty()
        || !current.footers.is_empty()
        || current.properties.page_width.is_some()
        || current.properties.page_height.is_some()
        || current.properties.orientation.is_some()
        || current.properties.margins.is_some()
        || current.properties.columns.is_some()
        || sections.is_empty()
    {
        sections.push(current);
    }

    Ok(sections)
}

fn parse_block_until(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    end_tag: &[u8],
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
                if e.name().as_ref() == end_tag {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(content)
}

fn parse_doc_part(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<GlossaryEntry, ParseError> {
    let mut entry = GlossaryEntry::new();
    entry.span = Some(SourceSpan::new("word/glossary/document.xml"));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:docPartPr" => {
                    let (name, gallery) = parse_doc_part_pr(reader)?;
                    entry.name = name;
                    entry.gallery = gallery;
                }
                b"w:docPartBody" => {
                    let content = parse_block_until(parser, reader, rels, b"w:docPartBody")?;
                    entry.content.extend(content);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:docPart" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/glossary/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(entry)
}

fn parse_doc_part_pr(
    reader: &mut Reader<&[u8]>,
) -> Result<(Option<String>, Option<String>), ParseError> {
    let mut name = None;
    let mut gallery = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:name" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        name = Some(val);
                    }
                }
                b"w:gallery" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        gallery = Some(val);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:docPartPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/glossary/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok((name, gallery))
}

fn apply_section_refs(
    reader: &mut Reader<&[u8]>,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<SectionRef, ParseError> {
    let mut headers = Vec::new();
    let mut footers = Vec::new();
    let mut properties = SectionProperties::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:headerReference" | b"w:footerReference" => {
                    if let Some(map) = header_footer_map {
                        if let Some(id) = attr_value(&e, b"r:id") {
                            if let Some(node_id) = map.get(&id) {
                                if e.name().as_ref() == b"w:headerReference" {
                                    headers.push(*node_id);
                                } else {
                                    footers.push(*node_id);
                                }
                            }
                        }
                    }
                }
                b"w:pgSz" => {
                    if let Some(val) = attr_value(&e, b"w:w").and_then(|v| v.parse().ok()) {
                        properties.page_width = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:h").and_then(|v| v.parse().ok()) {
                        properties.page_height = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:orient") {
                        properties.orientation = match val.as_str() {
                            "landscape" => Some(PageOrientation::Landscape),
                            "portrait" => Some(PageOrientation::Portrait),
                            _ => None,
                        };
                    }
                }
                b"w:pgMar" => {
                    let mut margins = properties.margins.take().unwrap_or(PageMargins {
                        top: 0,
                        bottom: 0,
                        left: 0,
                        right: 0,
                        header: None,
                        footer: None,
                        gutter: None,
                    });
                    if let Some(val) = attr_value(&e, b"w:top").and_then(|v| v.parse().ok()) {
                        margins.top = val;
                    }
                    if let Some(val) = attr_value(&e, b"w:bottom").and_then(|v| v.parse().ok()) {
                        margins.bottom = val;
                    }
                    if let Some(val) = attr_value(&e, b"w:left").and_then(|v| v.parse().ok()) {
                        margins.left = val;
                    }
                    if let Some(val) = attr_value(&e, b"w:right").and_then(|v| v.parse().ok()) {
                        margins.right = val;
                    }
                    if let Some(val) = attr_value(&e, b"w:header").and_then(|v| v.parse().ok()) {
                        margins.header = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:footer").and_then(|v| v.parse().ok()) {
                        margins.footer = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:gutter").and_then(|v| v.parse().ok()) {
                        margins.gutter = Some(val);
                    }
                    properties.margins = Some(margins);
                }
                b"w:cols" => {
                    if let Some(val) = attr_value(&e, b"w:num").and_then(|v| v.parse().ok()) {
                        properties.columns = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:space").and_then(|v| v.parse().ok()) {
                        properties.column_spacing = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:sep") {
                        properties.column_separator =
                            Some(val == "1" || val.eq_ignore_ascii_case("true"));
                    }
                }
                b"w:type" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        properties.section_type = match val.as_str() {
                            "continuous" => Some(SectionType::Continuous),
                            "evenPage" => Some(SectionType::EvenPage),
                            "oddPage" => Some(SectionType::OddPage),
                            "nextPage" => Some(SectionType::NextPage),
                            _ => None,
                        };
                    }
                }
                b"w:titlePg" => {
                    properties.title_page = Some(bool_from_val(&e));
                }
                b"w:pgNumType" => {
                    let mut numbering = properties.page_numbering.take().unwrap_or_default();
                    if let Some(val) = attr_value(&e, b"w:start").and_then(|v| v.parse().ok()) {
                        numbering.start = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:fmt") {
                        numbering.format = Some(val);
                    }
                    properties.page_numbering = Some(numbering);
                }
                b"w:lnNumType" | b"w:lineNumberType" => {
                    let mut numbering = properties.line_numbering.take().unwrap_or_default();
                    if let Some(val) = attr_value(&e, b"w:start").and_then(|v| v.parse().ok()) {
                        numbering.start = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:countBy").and_then(|v| v.parse().ok()) {
                        numbering.count_by = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:distance").and_then(|v| v.parse().ok()) {
                        numbering.distance = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:restart") {
                        numbering.restart = match val.as_str() {
                            "newPage" => Some(LineNumberRestart::NewPage),
                            "newSection" => Some(LineNumberRestart::NewSection),
                            "continuous" => Some(LineNumberRestart::Continuous),
                            _ => None,
                        };
                    }
                    properties.line_numbering = Some(numbering);
                }
                b"w:pgBorders" => {
                    if let Some(borders) = parse_page_borders(reader)? {
                        properties.page_borders = Some(borders);
                    }
                }
                b"w:textDirection" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        properties.text_direction = Some(val);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:sectPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(SectionRef {
        headers,
        footers,
        properties,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::relationships::Relationships;
    use quick_xml::events::Event;
    use quick_xml::Reader;

    #[test]
    fn test_parse_glossary_document() {
        let xml = r#"
        <w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:docParts>
            <w:docPart>
              <w:docPartPr>
                <w:name w:val="BlockOne"/>
                <w:gallery w:val="QuickParts"/>
              </w:docPartPr>
              <w:docPartBody>
                <w:p>
                  <w:r><w:t>Hello</w:t></w:r>
                </w:p>
              </w:docPartBody>
            </w:docPart>
          </w:docParts>
        </w:glossaryDocument>
        "#;

        let mut parser = DocxParser::new();
        let glossary_id = parser
            .parse_glossary_document(xml, &Relationships::default())
            .expect("parse glossary");
        let store = parser.into_store();
        let glossary = match store.get(glossary_id) {
            Some(docir_core::ir::IRNode::GlossaryDocument(doc)) => doc,
            _ => panic!("missing glossary"),
        };
        assert_eq!(glossary.entries.len(), 1);
        let entry_id = glossary.entries[0];
        let entry = match store.get(entry_id) {
            Some(docir_core::ir::IRNode::GlossaryEntry(e)) => e,
            _ => panic!("missing glossary entry"),
        };
        assert_eq!(entry.name.as_deref(), Some("BlockOne"));
        assert_eq!(entry.gallery.as_deref(), Some("QuickParts"));
        assert_eq!(entry.content.len(), 1);
    }

    #[test]
    fn test_parse_section_properties_extended() {
        let xml = r#"
        <w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pgSz w:w="12240" w:h="15840" w:orient="landscape"/>
          <w:pgMar w:top="720" w:bottom="720" w:left="720" w:right="720" w:gutter="180"/>
          <w:cols w:num="2" w:space="720" w:sep="1"/>
          <w:type w:val="continuous"/>
          <w:titlePg/>
          <w:pgNumType w:start="3" w:fmt="upperRoman"/>
          <w:lnNumType w:start="1" w:countBy="2" w:distance="240" w:restart="newPage"/>
          <w:pgBorders>
            <w:top w:val="single" w:sz="8" w:color="FF0000"/>
          </w:pgBorders>
          <w:textDirection w:val="tbRl"/>
        </w:sectPr>
        "#;

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:sectPr" => break,
                Ok(Event::Eof) => panic!("no sectPr"),
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }

        let section = apply_section_refs(&mut reader, None).expect("section");
        let props = section.properties;
        assert_eq!(props.page_width, Some(12240));
        assert_eq!(props.page_height, Some(15840));
        assert_eq!(props.orientation, Some(PageOrientation::Landscape));
        assert_eq!(props.columns, Some(2));
        assert_eq!(props.column_spacing, Some(720));
        assert_eq!(props.column_separator, Some(true));
        assert_eq!(props.section_type, Some(SectionType::Continuous));
        assert_eq!(props.title_page, Some(true));
        assert_eq!(props.text_direction.as_deref(), Some("tbRl"));
        assert_eq!(props.margins.as_ref().and_then(|m| m.gutter), Some(180));
        assert_eq!(props.page_numbering.as_ref().and_then(|n| n.start), Some(3));
        assert_eq!(
            props
                .page_numbering
                .as_ref()
                .and_then(|n| n.format.as_deref()),
            Some("upperRoman")
        );
        assert_eq!(
            props.line_numbering.as_ref().and_then(|n| n.count_by),
            Some(2)
        );
        assert_eq!(
            props.line_numbering.as_ref().and_then(|n| n.restart),
            Some(LineNumberRestart::NewPage)
        );
        assert!(props
            .page_borders
            .as_ref()
            .and_then(|b| b.top.as_ref())
            .is_some());
    }

    #[test]
    fn test_parse_styles_with_table_props() {
        let xml = r#"
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:styleId="MyStyle">
            <w:name w:val="My Style"/>
            <w:rPr>
              <w:b/>
              <w:u w:val="single"/>
              <w:color w:val="FF0000"/>
            </w:rPr>
            <w:pPr>
              <w:jc w:val="center"/>
              <w:spacing w:before="120"/>
            </w:pPr>
            <w:tblPr>
              <w:tblW w:w="5000" w:type="dxa"/>
            </w:tblPr>
          </w:style>
        </w:styles>
        "#;

        let mut parser = DocxParser::new();
        let styles_id = parser.parse_styles(xml).expect("styles");
        let store = parser.into_store();
        let styles = match store.get(styles_id) {
            Some(docir_core::ir::IRNode::StyleSet(s)) => s,
            _ => panic!("missing styles"),
        };
        let style = &styles.styles[0];
        assert_eq!(style.name.as_deref(), Some("My Style"));
        let run_props = style.run_props.as_ref().expect("run props");
        assert_eq!(run_props.bold, Some(true));
        assert_eq!(run_props.underline, Some(UnderlineStyle::Single));
        assert_eq!(run_props.color.as_deref(), Some("FF0000"));
        let para_props = style.paragraph_props.as_ref().expect("para props");
        assert_eq!(para_props.alignment, Some(TextAlignment::Center));
        assert_eq!(
            para_props.spacing.as_ref().and_then(|s| s.before),
            Some(120)
        );
        assert!(style
            .table_props
            .as_ref()
            .and_then(|t| t.width.as_ref())
            .is_some());
    }

    #[test]
    fn test_parse_numbering_level_props() {
        let xml = r#"
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="1">
            <w:lvl w:ilvl="0">
              <w:start w:val="1"/>
              <w:numFmt w:val="decimal"/>
              <w:lvlText w:val="%1."/>
              <w:lvlJc w:val="right"/>
              <w:suff w:val="space"/>
              <w:pPr>
                <w:spacing w:after="200"/>
              </w:pPr>
              <w:rPr>
                <w:i/>
              </w:rPr>
            </w:lvl>
          </w:abstractNum>
        </w:numbering>
        "#;

        let mut parser = DocxParser::new();
        let numbering_id = parser.parse_numbering(xml).expect("numbering");
        let store = parser.into_store();
        let numbering = match store.get(numbering_id) {
            Some(docir_core::ir::IRNode::NumberingSet(n)) => n,
            _ => panic!("missing numbering"),
        };
        let level = &numbering.abstract_nums[0].levels[0];
        assert_eq!(level.alignment, Some(TextAlignment::Right));
        assert_eq!(level.suffix.as_deref(), Some("space"));
        assert_eq!(
            level
                .paragraph_props
                .as_ref()
                .and_then(|p| p.spacing.as_ref())
                .and_then(|s| s.after),
            Some(200)
        );
        assert_eq!(level.run_props.as_ref().and_then(|r| r.italic), Some(true));
    }

    #[test]
    fn test_parse_drawing_smartart_targets() {
        let xml = r#"
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:drawing>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <dgm:relIds r:dm="rId1" r:lo="rId2" r:cs="rId3"/>
              </a:graphicData>
            </a:graphic>
          </w:drawing>
        </w:r>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
            Target="diagrams/data1.xml"/>
          <Relationship Id="rId2"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"
            Target="diagrams/layout1.xml"/>
          <Relationship Id="rId3"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"
            Target="diagrams/colors1.xml"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels");

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut parser = DocxParser::new();

        let mut buf = Vec::new();
        let mut run = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                    run = Some(parse_run(&mut parser, &mut reader, &rels).expect("run"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }

        let run = run.expect("run parsed");
        assert_eq!(run.embedded.len(), 1);
        let store = parser.into_store();
        let shape = match store.get(run.embedded[0]) {
            Some(docir_core::ir::IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };
        assert_eq!(shape.related_targets.len(), 3);
        assert!(shape
            .related_targets
            .contains(&"word/diagrams/data1.xml".to_string()));
        assert!(shape
            .related_targets
            .contains(&"word/diagrams/layout1.xml".to_string()));
        assert!(shape
            .related_targets
            .contains(&"word/diagrams/colors1.xml".to_string()));
    }

    #[test]
    fn test_parse_drawing_normalizes_targets() {
        let xml = r#"
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:drawing>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <a:blip r:embed="rIdImg"/>
                <dgm:relIds r:dm="rId1" r:lo="rId2"/>
              </a:graphicData>
            </a:graphic>
          </w:drawing>
        </w:r>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdImg"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="../media/image1.png"/>
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
            Target="../diagrams/data1.xml"/>
          <Relationship Id="rId2"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"
            Target="./diagrams/layout1.xml"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels");

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut parser = DocxParser::new();

        let mut buf = Vec::new();
        let mut run = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                    run = Some(parse_run(&mut parser, &mut reader, &rels).expect("run"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let run = run.expect("run parse");
        let store = parser.into_store();
        let shape = match store.get(run.embedded[0]) {
            Some(docir_core::ir::IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };
        assert_eq!(
            shape.media_target.as_deref(),
            Some("word/diagrams/data1.xml")
        );
        assert!(shape
            .related_targets
            .contains(&"word/diagrams/data1.xml".to_string()));
        assert!(shape
            .related_targets
            .contains(&"word/diagrams/layout1.xml".to_string()));
    }

    #[test]
    fn test_parse_drawing_text_and_hyperlink() {
        let xml = r#"
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:drawing>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <a:prstGeom prst="rect"/>
                <a:blip r:embed="rIdImg"/>
                <a:txBody>
                  <a:p>
                    <a:r>
                      <a:rPr b="1"/>
                      <a:t>Hello</a:t>
                    </a:r>
                  </a:p>
                </a:txBody>
                <a:hlinkClick r:id="rIdLink"/>
              </a:graphicData>
            </a:graphic>
          </w:drawing>
        </w:r>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdImg"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="media/image1.png"/>
          <Relationship Id="rIdLink"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            Target="https://example.com"
            TargetMode="External"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels");

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut parser = DocxParser::new();

        let mut buf = Vec::new();
        let mut run = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                    run = Some(parse_run(&mut parser, &mut reader, &rels).expect("run"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let run = run.expect("run parse");
        let store = parser.into_store();
        let shape = match store.get(run.embedded[0]) {
            Some(docir_core::ir::IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };
        assert_eq!(shape.shape_type, docir_core::ir::ShapeType::Rectangle);
        assert_eq!(shape.hyperlink.as_deref(), Some("https://example.com"));
        let text = shape.text.as_ref().expect("shape text");
        assert_eq!(text.paragraphs.len(), 1);
        assert_eq!(text.paragraphs[0].runs[0].text, "Hello");
    }

    #[test]
    fn test_parse_revisions_move_and_format() {
        let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:moveFrom w:author="Alice">
            <w:r><w:t>Old</w:t></w:r>
          </w:moveFrom>
          <w:moveTo w:author="Bob">
            <w:r><w:t>New</w:t></w:r>
          </w:moveTo>
          <w:rPrChange w:author="Carol">
            <w:rPr><w:b/></w:rPr>
          </w:rPrChange>
        </w:p>
        "#;

        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut para = None;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                    let parsed =
                        parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para");
                    para = Some(parsed);
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }

        let para = para.expect("para parsed");
        let store = parser.into_store();
        let mut types = Vec::new();
        let para_node = match store.get(para.id) {
            Some(docir_core::ir::IRNode::Paragraph(p)) => p,
            _ => panic!("missing paragraph"),
        };
        for id in &para_node.runs {
            if let Some(docir_core::ir::IRNode::Revision(rev)) = store.get(*id) {
                types.push(rev.change_type);
            }
        }
        types.sort_by_key(|t| format!("{t:?}"));
        assert!(types.contains(&RevisionType::MoveFrom));
        assert!(types.contains(&RevisionType::MoveTo));
        assert!(types.contains(&RevisionType::FormatChange));
    }

    #[test]
    fn test_parse_comments_and_notes_metadata() {
        let comments_xml = r#"
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:comment w:id="1" w:author="Alice" w:date="2020-01-01T00:00:00Z"
                     w:initials="AL" w:parentId="0" w:paraId="ABC" w:done="1">
            <w:p><w:r><w:t>Note</w:t></w:r></w:p>
          </w:comment>
        </w:comments>
        "#;
        let notes_xml = r#"
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:footnote w:id="2" w:type="separator">
            <w:p><w:r><w:t>---</w:t></w:r></w:p>
          </w:footnote>
        </w:footnotes>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let comment_ids = parser
            .parse_comments(comments_xml, &rels)
            .expect("comments");
        let note_ids = parser
            .parse_notes(notes_xml, NoteKind::Footnote, &rels)
            .expect("notes");
        let store = parser.into_store();

        let comment = match store.get(comment_ids[0]) {
            Some(docir_core::ir::IRNode::Comment(c)) => c,
            _ => panic!("missing comment"),
        };
        assert_eq!(comment.author.as_deref(), Some("Alice"));
        assert_eq!(comment.initials.as_deref(), Some("AL"));
        assert_eq!(comment.parent_id.as_deref(), Some("0"));
        assert_eq!(comment.para_id.as_deref(), Some("ABC"));
        assert_eq!(comment.done, Some(true));

        let footnote = match store.get(note_ids[0]) {
            Some(docir_core::ir::IRNode::Footnote(n)) => n,
            _ => panic!("missing footnote"),
        };
        assert_eq!(footnote.note_type.as_deref(), Some("separator"));
    }

    #[test]
    fn test_parse_bookmark_columns() {
        let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:bookmarkStart w:id="5" w:name="BM1" w:colFirst="1" w:colLast="3"/>
          <w:r><w:t>Text</w:t></w:r>
          <w:bookmarkEnd w:id="5"/>
        </w:p>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut para = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                    para =
                        Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let para = para.expect("para parsed");
        let store = parser.into_store();
        let para_node = match store.get(para.id) {
            Some(docir_core::ir::IRNode::Paragraph(p)) => p,
            _ => panic!("missing paragraph"),
        };
        let mut bookmark = None;
        for id in &para_node.runs {
            if let Some(docir_core::ir::IRNode::BookmarkStart(bm)) = store.get(*id) {
                bookmark = Some(bm);
                break;
            }
        }
        let bm = bookmark.expect("bookmark");
        assert_eq!(bm.name.as_deref(), Some("BM1"));
        assert_eq!(bm.col_first, Some(1));
        assert_eq!(bm.col_last, Some(3));
    }

    #[test]
    fn test_parse_field_instruction_hyperlink() {
        let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:fldSimple w:instr="HYPERLINK &quot;https://example.com&quot; \\t &quot;_blank&quot;">
            <w:r><w:t>Link</w:t></w:r>
          </w:fldSimple>
        </w:p>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut para = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                    para =
                        Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let para = para.expect("para parsed");
        let store = parser.into_store();
        let para_node = match store.get(para.id) {
            Some(docir_core::ir::IRNode::Paragraph(p)) => p,
            _ => panic!("missing paragraph"),
        };
        let mut field = None;
        for id in &para_node.runs {
            if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id) {
                field = Some(f);
                break;
            }
        }
        let field = field.expect("field");
        let parsed = field.instruction_parsed.as_ref().expect("parsed");
        assert!(matches!(parsed.kind, docir_core::ir::FieldKind::Hyperlink));
        assert!(parsed.args.iter().any(|a| a == "https://example.com"));
        assert!(parsed.switches.iter().any(|s| s == "\\t"));
    }

    #[test]
    fn test_parse_field_instruction_includetext() {
        let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:fldSimple w:instr="INCLUDETEXT &quot;C:\\docs\\file.docx&quot; \\m">
            <w:r><w:t>Include</w:t></w:r>
          </w:fldSimple>
        </w:p>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut para = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                    para =
                        Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let para = para.expect("para parsed");
        let store = parser.into_store();
        let para_node = match store.get(para.id) {
            Some(docir_core::ir::IRNode::Paragraph(p)) => p,
            _ => panic!("missing paragraph"),
        };
        let mut field = None;
        for id in &para_node.runs {
            if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id) {
                field = Some(f);
                break;
            }
        }
        let field = field.expect("field");
        let parsed = field.instruction_parsed.as_ref().expect("parsed");
        assert!(matches!(
            parsed.kind,
            docir_core::ir::FieldKind::IncludeText
        ));
        assert!(parsed
            .args
            .iter()
            .any(|a| a.contains("C:") && a.contains("docs") && a.contains("file.docx")));
        assert!(parsed.switches.iter().any(|s| s == "\\m"));
    }

    #[test]
    fn test_parse_field_instruction_mergefield() {
        let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:fldSimple w:instr="MERGEFIELD  CustomerName  \\* MERGEFORMAT">
            <w:r><w:t>Value</w:t></w:r>
          </w:fldSimple>
        </w:p>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut para = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                    para =
                        Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let para = para.expect("para parsed");
        let store = parser.into_store();
        let para_node = match store.get(para.id) {
            Some(docir_core::ir::IRNode::Paragraph(p)) => p,
            _ => panic!("missing paragraph"),
        };
        let mut field = None;
        for id in &para_node.runs {
            if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id) {
                field = Some(f);
                break;
            }
        }
        let field = field.expect("field");
        let parsed = field.instruction_parsed.as_ref().expect("parsed");
        assert!(matches!(parsed.kind, docir_core::ir::FieldKind::MergeField));
        assert!(parsed.args.iter().any(|a| a == "CustomerName"));
        assert!(parsed.switches.iter().any(|s| s == "\\*"));
    }

    #[test]
    fn test_parse_field_instruction_date() {
        let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:fldSimple w:instr="DATE \\@ &quot;MMMM d, yyyy&quot;">
            <w:r><w:t>Today</w:t></w:r>
          </w:fldSimple>
        </w:p>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut para = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                    para =
                        Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let para = para.expect("para parsed");
        let store = parser.into_store();
        let para_node = match store.get(para.id) {
            Some(docir_core::ir::IRNode::Paragraph(p)) => p,
            _ => panic!("missing paragraph"),
        };
        let mut field = None;
        for id in &para_node.runs {
            if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id) {
                field = Some(f);
                break;
            }
        }
        let field = field.expect("field");
        let parsed = field.instruction_parsed.as_ref().expect("parsed");
        assert!(matches!(parsed.kind, docir_core::ir::FieldKind::Date));
        assert!(parsed.switches.iter().any(|s| s == "\\@"));
    }

    #[test]
    fn test_parse_field_instruction_ref_pageref() {
        let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:fldSimple w:instr="REF Bookmark1 \\h">
            <w:r><w:t>Ref</w:t></w:r>
          </w:fldSimple>
          <w:fldSimple w:instr="PAGEREF Bookmark1 \\p">
            <w:r><w:t>PageRef</w:t></w:r>
          </w:fldSimple>
        </w:p>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut para = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                    para =
                        Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let para = para.expect("para parsed");
        let store = parser.into_store();
        let para_node = match store.get(para.id) {
            Some(docir_core::ir::IRNode::Paragraph(p)) => p,
            _ => panic!("missing paragraph"),
        };
        let mut kinds = Vec::new();
        for id in &para_node.runs {
            if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id) {
                if let Some(parsed) = f.instruction_parsed.as_ref() {
                    kinds.push(parsed.kind.clone());
                }
            }
        }
        assert!(kinds
            .iter()
            .any(|k| matches!(k, docir_core::ir::FieldKind::Ref)));
        assert!(kinds
            .iter()
            .any(|k| matches!(k, docir_core::ir::FieldKind::PageRef)));
    }

    #[test]
    fn test_parse_field_instruction_extended() {
        let parsed = parse_field_instruction("DDE \"cmd\" \"args\"").expect("parsed");
        assert!(matches!(parsed.kind, docir_core::ir::FieldKind::Dde));

        let parsed = parse_field_instruction("DDEAUTO \"cmd\" \"args\"").expect("parsed");
        assert!(matches!(parsed.kind, docir_core::ir::FieldKind::DdeAuto));

        let parsed = parse_field_instruction("AUTOTEXT MyEntry").expect("parsed");
        assert!(matches!(parsed.kind, docir_core::ir::FieldKind::AutoText));

        let parsed = parse_field_instruction("AUTOCORRECT MyEntry").expect("parsed");
        assert!(matches!(
            parsed.kind,
            docir_core::ir::FieldKind::AutoCorrect
        ));

        let parsed = parse_field_instruction("INCLUDEPICTURE \"image.png\"").expect("parsed");
        assert!(matches!(
            parsed.kind,
            docir_core::ir::FieldKind::IncludePicture
        ));
    }

    #[test]
    fn test_parse_hyperlink_anchor_tooltip() {
        let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:hyperlink w:anchor="BM1" w:tooltip="Go to bookmark">
            <w:r><w:t>Link</w:t></w:r>
          </w:hyperlink>
        </w:p>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut para = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                    para =
                        Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let para = para.expect("para parsed");
        let store = parser.into_store();
        let para_node = match store.get(para.id) {
            Some(docir_core::ir::IRNode::Paragraph(p)) => p,
            _ => panic!("missing paragraph"),
        };
        let mut link = None;
        for id in &para_node.runs {
            if let Some(docir_core::ir::IRNode::Hyperlink(h)) = store.get(*id) {
                link = Some(h);
                break;
            }
        }
        let link = link.expect("hyperlink");
        assert_eq!(link.target, "#BM1");
        assert_eq!(link.tooltip.as_deref(), Some("Go to bookmark"));
        assert_eq!(link.is_external, false);
    }

    #[test]
    fn test_parse_content_control_data_binding() {
        let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:sdt>
            <w:sdtPr>
              <w:tag w:val="customer"/>
              <w:dataBinding w:xpath="/customer/name" w:storeItemID="{1234}"
                             w:prefixMappings="xmlns:c='urn:customer'"/>
            </w:sdtPr>
            <w:sdtContent>
              <w:r><w:t>Value</w:t></w:r>
            </w:sdtContent>
          </w:sdt>
        </w:p>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut para = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                    para =
                        Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let para = para.expect("para parsed");
        let store = parser.into_store();
        let para_node = match store.get(para.id) {
            Some(docir_core::ir::IRNode::Paragraph(p)) => p,
            _ => panic!("missing paragraph"),
        };
        let mut control = None;
        for id in &para_node.runs {
            if let Some(docir_core::ir::IRNode::ContentControl(c)) = store.get(*id) {
                control = Some(c);
                break;
            }
        }
        let control = control.expect("content control");
        assert_eq!(control.tag.as_deref(), Some("customer"));
        assert_eq!(
            control.data_binding_xpath.as_deref(),
            Some("/customer/name")
        );
        assert_eq!(
            control.data_binding_store_item_id.as_deref(),
            Some("{1234}")
        );
        assert_eq!(
            control.data_binding_prefix_mappings.as_deref(),
            Some("xmlns:c='urn:customer'")
        );
    }

    #[test]
    fn test_parse_table_grid_and_properties() {
        let xml = r#"
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tblPr>
            <w:tblW w:w="5000" w:type="dxa"/>
            <w:jc w:val="center"/>
            <w:tblStyle w:val="TableStyle1"/>
            <w:tblBorders>
              <w:top w:val="single" w:sz="8" w:color="FF0000"/>
            </w:tblBorders>
            <w:tblCellMar>
              <w:top w:w="100"/>
              <w:left w:w="120"/>
            </w:tblCellMar>
          </w:tblPr>
          <w:tblGrid>
            <w:gridCol w:w="2400"/>
            <w:gridCol w:w="2600"/>
          </w:tblGrid>
          <w:tr>
            <w:trPr>
              <w:trHeight w:val="300" w:hRule="exact"/>
              <w:tblHeader/>
              <w:cantSplit w:val="1"/>
            </w:trPr>
            <w:tc>
              <w:tcPr>
                <w:shd w:fill="FFFF00"/>
              </w:tcPr>
              <w:p><w:r><w:t>A</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut table_id = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:tbl" => {
                    table_id = Some(parse_table(&mut parser, &mut reader, &rels).expect("table"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let table_id = table_id.expect("table parsed");
        let store = parser.into_store();
        let table = match store.get(table_id) {
            Some(docir_core::ir::IRNode::Table(t)) => t,
            _ => panic!("missing table"),
        };
        assert_eq!(table.grid.len(), 2);
        assert_eq!(table.grid[0].width, 2400);
        assert_eq!(table.grid[1].width, 2600);
        let props = &table.properties;
        assert_eq!(props.width.as_ref().map(|w| w.value), Some(5000));
        assert!(matches!(
            props.alignment,
            Some(docir_core::ir::TableAlignment::Center)
        ));
        assert_eq!(props.style_id.as_deref(), Some("TableStyle1"));
        assert_eq!(props.cell_margins.as_ref().and_then(|m| m.top), Some(100));
        assert_eq!(props.cell_margins.as_ref().and_then(|m| m.left), Some(120));
        assert!(props
            .borders
            .as_ref()
            .and_then(|b| b.top.as_ref())
            .is_some());

        let row = match store.get(table.rows[0]) {
            Some(docir_core::ir::IRNode::TableRow(r)) => r,
            _ => panic!("missing row"),
        };
        assert_eq!(row.properties.height.as_ref().map(|h| h.value), Some(300));
        assert!(matches!(
            row.properties.height.as_ref().map(|h| h.rule),
            Some(docir_core::ir::RowHeightRule::Exact)
        ));
        assert_eq!(row.properties.is_header, Some(true));
        assert_eq!(row.properties.cant_split, Some(true));

        let cell = match store.get(row.cells[0]) {
            Some(docir_core::ir::IRNode::TableCell(c)) => c,
            _ => panic!("missing cell"),
        };
        assert_eq!(cell.properties.shading.as_deref(), Some("FFFF00"));
    }

    #[test]
    fn test_parse_run_properties_caps_and_style() {
        let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r>
            <w:rPr>
              <w:rStyle w:val="Emphasis"/>
              <w:caps/>
              <w:smallCaps w:val="0"/>
            </w:rPr>
            <w:t>Text</w:t>
          </w:r>
        </w:p>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut para = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                    para =
                        Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let para = para.expect("para parsed");
        let store = parser.into_store();
        let para_node = match store.get(para.id) {
            Some(docir_core::ir::IRNode::Paragraph(p)) => p,
            _ => panic!("missing paragraph"),
        };
        let run = match store.get(para_node.runs[0]) {
            Some(docir_core::ir::IRNode::Run(r)) => r,
            _ => panic!("missing run"),
        };
        assert_eq!(run.properties.style_id.as_deref(), Some("Emphasis"));
        assert_eq!(run.properties.all_caps, Some(true));
        assert_eq!(run.properties.small_caps, Some(false));
    }

    #[test]
    fn test_parse_paragraph_borders_and_flags() {
        let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pPr>
            <w:keepNext/>
            <w:keepLines w:val="1"/>
            <w:pageBreakBefore/>
            <w:widowControl w:val="0"/>
            <w:pBdr>
              <w:top w:val="single" w:sz="4" w:color="00FF00"/>
            </w:pBdr>
          </w:pPr>
          <w:r><w:t>Para</w:t></w:r>
        </w:p>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut para = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                    para =
                        Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let para = para.expect("para parsed");
        let store = parser.into_store();
        let para_node = match store.get(para.id) {
            Some(docir_core::ir::IRNode::Paragraph(p)) => p,
            _ => panic!("missing paragraph"),
        };
        assert_eq!(para_node.properties.keep_next, Some(true));
        assert_eq!(para_node.properties.keep_lines, Some(true));
        assert_eq!(para_node.properties.page_break_before, Some(true));
        assert_eq!(para_node.properties.widow_control, Some(false));
        let border = para_node
            .properties
            .borders
            .as_ref()
            .and_then(|b| b.top.as_ref());
        assert!(border.is_some());
    }

    #[test]
    fn test_parse_note_references_as_fields() {
        let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:footnoteReference w:id="1"/></w:r>
          <w:r><w:endnoteReference w:id="2"/></w:r>
        </w:p>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut para = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                    para =
                        Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {}", e),
                _ => {}
            }
            buf.clear();
        }
        let para = para.expect("para parsed");
        let store = parser.into_store();
        let para_node = match store.get(para.id) {
            Some(docir_core::ir::IRNode::Paragraph(p)) => p,
            _ => panic!("missing paragraph"),
        };
        let mut kinds = Vec::new();
        let mut args = Vec::new();
        for id in &para_node.runs {
            if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id) {
                if let Some(parsed) = f.instruction_parsed.as_ref() {
                    kinds.push(parsed.kind.clone());
                    args.extend(parsed.args.clone());
                }
            }
        }
        assert!(kinds
            .iter()
            .any(|k| matches!(k, docir_core::ir::FieldKind::FootnoteRef)));
        assert!(kinds
            .iter()
            .any(|k| matches!(k, docir_core::ir::FieldKind::EndnoteRef)));
        assert!(args.iter().any(|a| a == "1"));
        assert!(args.iter().any(|a| a == "2"));
    }
}

fn parse_paragraph(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<ParagraphParse, ParseError> {
    let mut para = Paragraph::new();
    let mut field_active = false;
    let mut field_instr = String::new();
    let mut field_runs: Vec<NodeId> = Vec::new();
    let mut field_instr_done = false;
    let mut section_ref: Option<SectionRef> = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:pPr" => {
                    section_ref = parse_paragraph_properties(reader, &mut para, header_footer_map)?;
                }
                b"w:r" => {
                    let run = parse_run(parser, reader, rels)?;
                    let run_id = run.run_id;
                    para.runs.push(run_id);
                    for emb in run.embedded {
                        para.runs.push(emb);
                    }
                    if field_active {
                        field_runs.push(run_id);
                        if run.has_instr && !field_instr_done {
                            field_instr.push_str(&run.text);
                        }
                    }
                    if let Some(char_type) = run.field_char.as_deref() {
                        match char_type {
                            "begin" => {
                                field_active = true;
                                field_instr_done = false;
                                field_instr.clear();
                                field_runs.clear();
                            }
                            "separate" => {
                                field_instr_done = true;
                            }
                            "end" => {
                                if field_active {
                                    let instr = if field_instr.trim().is_empty() {
                                        None
                                    } else {
                                        Some(field_instr.trim().to_string())
                                    };
                                    let mut field = Field::new(instr);
                                    field.runs = field_runs.clone();
                                    let field_id = field.id;
                                    parser.store.insert(docir_core::ir::IRNode::Field(field));
                                    para.runs.push(field_id);
                                }
                                field_active = false;
                                field_instr_done = false;
                                field_instr.clear();
                                field_runs.clear();
                            }
                            _ => {}
                        }
                    }
                }
                b"w:hyperlink" => {
                    let link_id = parse_hyperlink(parser, reader, rels, &e)?;
                    para.runs.push(link_id);
                }
                b"w:sdt" => {
                    let sdt_id = parse_sdt(parser, reader, rels, SdtMode::Inline)?;
                    para.runs.push(sdt_id);
                }
                b"w:fldSimple" => {
                    let instr = attr_value(&e, b"w:instr");
                    let field_id = parse_field(parser, reader, instr)?;
                    para.runs.push(field_id);
                }
                b"w:commentRangeStart" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentRangeStart::new(cid);
                        node.span = Some(span_from_reader(reader, "word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentRangeStart(node));
                        para.runs.push(node_id);
                    }
                }
                b"w:commentRangeEnd" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentRangeEnd::new(cid);
                        node.span = Some(span_from_reader(reader, "word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentRangeEnd(node));
                        para.runs.push(node_id);
                    }
                }
                b"w:commentReference" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentReference::new(cid);
                        node.span = Some(span_from_reader(reader, "word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentReference(node));
                        para.runs.push(node_id);
                    }
                }
                b"w:ins" => {
                    let rev_id =
                        parse_revision_inline(parser, reader, rels, &e, RevisionType::Insert)?;
                    para.runs.push(rev_id);
                }
                b"w:del" => {
                    let rev_id =
                        parse_revision_inline(parser, reader, rels, &e, RevisionType::Delete)?;
                    para.runs.push(rev_id);
                }
                b"w:moveFrom" => {
                    let rev_id =
                        parse_revision_inline(parser, reader, rels, &e, RevisionType::MoveFrom)?;
                    para.runs.push(rev_id);
                }
                b"w:moveTo" => {
                    let rev_id =
                        parse_revision_inline(parser, reader, rels, &e, RevisionType::MoveTo)?;
                    para.runs.push(rev_id);
                }
                b"w:pPrChange" | b"w:rPrChange" => {
                    let rev_id = parse_revision_inline(
                        parser,
                        reader,
                        rels,
                        &e,
                        RevisionType::FormatChange,
                    )?;
                    para.runs.push(rev_id);
                }
                b"w:bookmarkStart" => {
                    if let Some(bm_id) = attr_value(&e, b"w:id") {
                        let mut bm = docir_core::ir::BookmarkStart::new(bm_id);
                        bm.name = attr_value(&e, b"w:name");
                        bm.col_first = attr_value(&e, b"w:colFirst").and_then(|v| v.parse().ok());
                        bm.col_last = attr_value(&e, b"w:colLast").and_then(|v| v.parse().ok());
                        let bm_id = bm.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::BookmarkStart(bm));
                        para.runs.push(bm_id);
                    }
                }
                b"w:bookmarkEnd" => {
                    if let Some(bm_id) = attr_value(&e, b"w:id") {
                        let bm = docir_core::ir::BookmarkEnd::new(bm_id);
                        let bm_id = bm.id;
                        parser.store.insert(docir_core::ir::IRNode::BookmarkEnd(bm));
                        para.runs.push(bm_id);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:commentRangeStart" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentRangeStart::new(cid);
                        node.span = Some(span_from_reader(reader, "word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentRangeStart(node));
                        para.runs.push(node_id);
                    }
                }
                b"w:commentRangeEnd" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentRangeEnd::new(cid);
                        node.span = Some(span_from_reader(reader, "word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentRangeEnd(node));
                        para.runs.push(node_id);
                    }
                }
                b"w:commentReference" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentReference::new(cid);
                        node.span = Some(span_from_reader(reader, "word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentReference(node));
                        para.runs.push(node_id);
                    }
                }
                b"w:bookmarkStart" => {
                    if let Some(bm_id) = attr_value(&e, b"w:id") {
                        let mut bm = docir_core::ir::BookmarkStart::new(bm_id);
                        bm.name = attr_value(&e, b"w:name");
                        bm.col_first = attr_value(&e, b"w:colFirst").and_then(|v| v.parse().ok());
                        bm.col_last = attr_value(&e, b"w:colLast").and_then(|v| v.parse().ok());
                        let bm_id = bm.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::BookmarkStart(bm));
                        para.runs.push(bm_id);
                    }
                }
                b"w:bookmarkEnd" => {
                    if let Some(bm_id) = attr_value(&e, b"w:id") {
                        let bm = docir_core::ir::BookmarkEnd::new(bm_id);
                        let bm_id = bm.id;
                        parser.store.insert(docir_core::ir::IRNode::BookmarkEnd(bm));
                        para.runs.push(bm_id);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:p" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    para.span = Some(span_from_reader(reader, "word/document.xml"));
    let id = para.id;
    parser.store.insert(docir_core::ir::IRNode::Paragraph(para));
    Ok(ParagraphParse { id, section_ref })
}

fn parse_paragraph_simple(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<NodeId, ParseError> {
    Ok(parse_paragraph(parser, reader, rels, None)?.id)
}

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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    let id = field.id;
    parser.store.insert(docir_core::ir::IRNode::Field(field));
    Ok(id)
}

fn parse_field_instruction(instr: &str) -> Option<docir_core::ir::FieldInstruction> {
    let decoded = unescape_xml_entities(instr);
    let tokens = tokenize_field_instruction(&decoded);
    if tokens.is_empty() {
        return None;
    }
    let kind = match tokens[0].as_str() {
        "HYPERLINK" => docir_core::ir::FieldKind::Hyperlink,
        "INCLUDETEXT" => docir_core::ir::FieldKind::IncludeText,
        "INCLUDEPICTURE" => docir_core::ir::FieldKind::IncludePicture,
        "MERGEFIELD" => docir_core::ir::FieldKind::MergeField,
        "DATE" => docir_core::ir::FieldKind::Date,
        "REF" => docir_core::ir::FieldKind::Ref,
        "PAGEREF" => docir_core::ir::FieldKind::PageRef,
        "DDE" => docir_core::ir::FieldKind::Dde,
        "DDEAUTO" => docir_core::ir::FieldKind::DdeAuto,
        "AUTOTEXT" => docir_core::ir::FieldKind::AutoText,
        "AUTOCORRECT" => docir_core::ir::FieldKind::AutoCorrect,
        _ => docir_core::ir::FieldKind::Unknown,
    };
    let mut args = Vec::new();
    let mut switches = Vec::new();
    for tok in tokens.into_iter().skip(1) {
        if tok.starts_with('\\') {
            switches.push(normalize_switch(&tok));
        } else {
            args.push(tok);
        }
    }
    for sw in extract_switches(&decoded) {
        let sw = normalize_switch(&sw);
        if !switches.contains(&sw) {
            switches.push(sw);
        }
    }
    if decoded.contains('\t') && !switches.iter().any(|s| s == "\\t") {
        switches.push("\\t".to_string());
    }
    if matches!(kind, docir_core::ir::FieldKind::Hyperlink) {
        let mut normalized_args = Vec::new();
        for arg in args {
            if arg.len() == 1 && arg.chars().all(|c| c.is_ascii_alphabetic()) {
                let sw = format!("\\{arg}");
                if !switches.contains(&sw) {
                    switches.push(sw);
                }
            } else {
                normalized_args.push(arg);
            }
        }
        return Some(docir_core::ir::FieldInstruction {
            kind,
            args: normalized_args,
            switches,
        });
    }
    Some(docir_core::ir::FieldInstruction {
        kind,
        args,
        switches,
    })
}

fn tokenize_field_instruction(instr: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    let mut current = String::new();
    let mut in_quotes = false;
    let mut chars = instr.chars().peekable();
    while let Some(ch) = chars.next() {
        match ch {
            '"' => {
                in_quotes = !in_quotes;
            }
            ' ' | '\t' | '\r' | '\n' if !in_quotes => {
                if !current.is_empty() {
                    tokens.push(current.clone());
                    current.clear();
                }
                while matches!(chars.peek(), Some(' ' | '\t' | '\r' | '\n')) {
                    chars.next();
                }
            }
            _ => current.push(ch),
        }
    }
    if !current.is_empty() {
        tokens.push(current);
    }
    tokens
}

fn extract_switches(instr: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut in_quotes = false;
    let chars: Vec<char> = instr.chars().collect();
    let mut i = 0usize;
    while i < chars.len() {
        let ch = chars[i];
        if ch == '"' {
            in_quotes = !in_quotes;
            i += 1;
            continue;
        }
        if !in_quotes && ch == '\\' {
            let mut j = i + 1;
            if j >= chars.len() || chars[j].is_whitespace() {
                i += 1;
                continue;
            }
            let mut token = String::new();
            token.push('\\');
            while j < chars.len() && !chars[j].is_whitespace() {
                token.push(chars[j]);
                j += 1;
            }
            out.push(token);
            i = j;
            continue;
        }
        i += 1;
    }
    out
}

fn unescape_xml_entities(value: &str) -> String {
    value
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
}

fn normalize_switch(value: &str) -> String {
    let mut out = value.to_string();
    while out.starts_with("\\\\") {
        out.remove(0);
    }
    out
}

fn parse_table(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<NodeId, ParseError> {
    let mut table = Table::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:tblPr" => {
                    parse_table_properties(reader, &mut table.properties)?;
                }
                b"w:tblGrid" => {
                    table.grid = parse_table_grid(reader)?;
                }
                b"w:tr" => {
                    let row_id = parse_table_row(parser, reader, rels)?;
                    table.rows.push(row_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:tbl" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    table.span = Some(span_from_reader(reader, "word/document.xml"));
    let id = table.id;
    parser.store.insert(docir_core::ir::IRNode::Table(table));
    Ok(id)
}

fn parse_table_row(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<NodeId, ParseError> {
    let mut row = TableRow::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:trPr" => {
                    parse_table_row_properties(reader, &mut row.properties)?;
                }
                b"w:tc" => {
                    let cell_id = parse_table_cell(parser, reader, rels)?;
                    row.cells.push(cell_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:tr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    row.span = Some(span_from_reader(reader, "word/document.xml"));
    let id = row.id;
    parser.store.insert(docir_core::ir::IRNode::TableRow(row));
    Ok(id)
}

fn parse_table_cell(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<NodeId, ParseError> {
    let mut cell = TableCell::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:tcPr" => {
                    parse_table_cell_properties(reader, &mut cell.properties)?;
                }
                b"w:p" => {
                    let para_id = parse_paragraph_simple(parser, reader, rels)?;
                    cell.content.push(para_id);
                }
                b"w:tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    cell.content.push(table_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:tc" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    cell.span = Some(span_from_reader(reader, "word/document.xml"));
    let id = cell.id;
    parser.store.insert(docir_core::ir::IRNode::TableCell(cell));
    Ok(id)
}

fn parse_paragraph_properties(
    reader: &mut Reader<&[u8]>,
    para: &mut Paragraph,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<Option<SectionRef>, ParseError> {
    let mut buf = Vec::new();
    let mut section_ref: Option<SectionRef> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:pStyle" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        para.style_id = Some(val);
                    }
                }
                b"w:keepNext" => {
                    para.properties.keep_next = Some(bool_from_val(&e));
                }
                b"w:keepLines" => {
                    para.properties.keep_lines = Some(bool_from_val(&e));
                }
                b"w:pageBreakBefore" => {
                    para.properties.page_break_before = Some(bool_from_val(&e));
                }
                b"w:widowControl" => {
                    para.properties.widow_control = Some(bool_from_val(&e));
                }
                b"w:jc" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        para.properties.alignment = match val.as_str() {
                            "center" => Some(TextAlignment::Center),
                            "right" => Some(TextAlignment::Right),
                            "both" => Some(TextAlignment::Justify),
                            "distribute" => Some(TextAlignment::Distribute),
                            _ => Some(TextAlignment::Left),
                        };
                    }
                }
                b"w:ind" => {
                    let mut indent = para.properties.indentation.clone().unwrap_or_default();
                    if let Some(val) = attr_value(&e, b"w:left").and_then(|v| v.parse().ok()) {
                        indent.left = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:right").and_then(|v| v.parse().ok()) {
                        indent.right = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:firstLine").and_then(|v| v.parse().ok()) {
                        indent.first_line = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:hanging").and_then(|v| v.parse().ok()) {
                        indent.hanging = Some(val);
                    }
                    para.properties.indentation = Some(indent);
                }
                b"w:spacing" => {
                    let mut spacing = para.properties.spacing.clone().unwrap_or_default();
                    if let Some(val) = attr_value(&e, b"w:before").and_then(|v| v.parse().ok()) {
                        spacing.before = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:after").and_then(|v| v.parse().ok()) {
                        spacing.after = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:line").and_then(|v| v.parse().ok()) {
                        spacing.line = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:lineRule") {
                        spacing.line_rule = match val.as_str() {
                            "auto" => Some(LineSpacingRule::Auto),
                            "exact" => Some(LineSpacingRule::Exact),
                            "atLeast" => Some(LineSpacingRule::AtLeast),
                            _ => None,
                        };
                    }
                    para.properties.spacing = Some(spacing);
                }
                b"w:pBdr" => {
                    if let Some(borders) = parse_paragraph_borders(reader)? {
                        para.properties.borders = Some(borders);
                    }
                }
                b"w:outlineLvl" => {
                    if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                        para.properties.outline_level = Some(val);
                    }
                }
                b"w:numPr" => {
                    parse_numbering(reader, &mut para.properties)?;
                }
                b"w:sectPr" => {
                    section_ref = Some(apply_section_refs(reader, header_footer_map)?);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:pPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(section_ref)
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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_table_cell_properties(
    reader: &mut Reader<&[u8]>,
    props: &mut TableCellProperties,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:tcW" => {
                    if let Some(val) = attr_value(&e, b"w:w").and_then(|v| v.parse().ok()) {
                        let width_type = match attr_value(&e, b"w:type").as_deref() {
                            Some("dxa") => TableWidthType::Dxa,
                            Some("pct") => TableWidthType::Pct,
                            Some("auto") => TableWidthType::Auto,
                            _ => TableWidthType::Nil,
                        };
                        props.width = Some(TableWidth {
                            value: val,
                            width_type,
                        });
                    }
                }
                b"w:gridSpan" => {
                    if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                        props.grid_span = Some(val);
                    }
                }
                b"w:vMerge" => {
                    let merge = match attr_value(&e, b"w:val").as_deref() {
                        Some("restart") => MergeType::Restart,
                        Some("continue") => MergeType::Continue,
                        _ => MergeType::Continue,
                    };
                    props.vertical_merge = Some(merge);
                }
                b"w:vAlign" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.vertical_align = match val.as_str() {
                            "center" => Some(CellVerticalAlignment::Center),
                            "bottom" => Some(CellVerticalAlignment::Bottom),
                            _ => Some(CellVerticalAlignment::Top),
                        };
                    }
                }
                b"w:tcBorders" => {
                    if let Some(borders) = parse_table_borders(reader, b"w:tcBorders")? {
                        props.borders = Some(borders);
                    }
                }
                b"w:shd" => {
                    if let Some(fill) = attr_value(&e, b"w:fill") {
                        props.shading = Some(fill);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:tcW" => {
                    if let Some(val) = attr_value(&e, b"w:w").and_then(|v| v.parse().ok()) {
                        let width_type = match attr_value(&e, b"w:type").as_deref() {
                            Some("dxa") => TableWidthType::Dxa,
                            Some("pct") => TableWidthType::Pct,
                            Some("auto") => TableWidthType::Auto,
                            _ => TableWidthType::Nil,
                        };
                        props.width = Some(TableWidth {
                            value: val,
                            width_type,
                        });
                    }
                }
                b"w:gridSpan" => {
                    if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                        props.grid_span = Some(val);
                    }
                }
                b"w:vMerge" => {
                    let merge = match attr_value(&e, b"w:val").as_deref() {
                        Some("restart") => MergeType::Restart,
                        Some("continue") => MergeType::Continue,
                        _ => MergeType::Continue,
                    };
                    props.vertical_merge = Some(merge);
                }
                b"w:vAlign" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.vertical_align = match val.as_str() {
                            "center" => Some(CellVerticalAlignment::Center),
                            "bottom" => Some(CellVerticalAlignment::Bottom),
                            _ => Some(CellVerticalAlignment::Top),
                        };
                    }
                }
                b"w:shd" => {
                    if let Some(fill) = attr_value(&e, b"w:fill") {
                        props.shading = Some(fill);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:tcPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_table_properties(
    reader: &mut Reader<&[u8]>,
    props: &mut docir_core::ir::TableProperties,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:tblW" => {
                    if let Some(val) = attr_value(&e, b"w:w").and_then(|v| v.parse().ok()) {
                        let width_type = match attr_value(&e, b"w:type").as_deref() {
                            Some("dxa") => TableWidthType::Dxa,
                            Some("pct") => TableWidthType::Pct,
                            Some("auto") => TableWidthType::Auto,
                            _ => TableWidthType::Nil,
                        };
                        props.width = Some(TableWidth {
                            value: val,
                            width_type,
                        });
                    }
                }
                b"w:jc" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.alignment = match val.as_str() {
                            "center" => Some(TableAlignment::Center),
                            "right" => Some(TableAlignment::Right),
                            _ => Some(TableAlignment::Left),
                        };
                    }
                }
                b"w:tblStyle" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.style_id = Some(val);
                    }
                }
                b"w:tblBorders" => {
                    if let Some(borders) = parse_table_borders(reader, b"w:tblBorders")? {
                        props.borders = Some(borders);
                    }
                }
                b"w:tblCellMar" => {
                    if let Some(margins) = parse_cell_margins(reader)? {
                        props.cell_margins = Some(margins);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:tblW" => {
                    if let Some(val) = attr_value(&e, b"w:w").and_then(|v| v.parse().ok()) {
                        let width_type = match attr_value(&e, b"w:type").as_deref() {
                            Some("dxa") => TableWidthType::Dxa,
                            Some("pct") => TableWidthType::Pct,
                            Some("auto") => TableWidthType::Auto,
                            _ => TableWidthType::Nil,
                        };
                        props.width = Some(TableWidth {
                            value: val,
                            width_type,
                        });
                    }
                }
                b"w:jc" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.alignment = match val.as_str() {
                            "center" => Some(TableAlignment::Center),
                            "right" => Some(TableAlignment::Right),
                            _ => Some(TableAlignment::Left),
                        };
                    }
                }
                b"w:tblStyle" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.style_id = Some(val);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:tblPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_table_grid(
    reader: &mut Reader<&[u8]>,
) -> Result<Vec<docir_core::ir::GridColumn>, ParseError> {
    let mut buf = Vec::new();
    let mut grid = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"w:gridCol" {
                    if let Some(val) = attr_value(&e, b"w:w").and_then(|v| v.parse().ok()) {
                        grid.push(docir_core::ir::GridColumn { width: val });
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:tblGrid" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(grid)
}

fn parse_table_row_properties(
    reader: &mut Reader<&[u8]>,
    props: &mut docir_core::ir::TableRowProperties,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:trHeight" => {
                    if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                        let rule = match attr_value(&e, b"w:hRule").as_deref() {
                            Some("exact") => RowHeightRule::Exact,
                            Some("atLeast") => RowHeightRule::AtLeast,
                            _ => RowHeightRule::Auto,
                        };
                        props.height = Some(RowHeight { value: val, rule });
                    }
                }
                b"w:tblHeader" => {
                    let is_header = match attr_value(&e, b"w:val").as_deref() {
                        Some("0") | Some("false") => false,
                        _ => true,
                    };
                    props.is_header = Some(is_header);
                }
                b"w:cantSplit" => {
                    let cant_split = match attr_value(&e, b"w:val").as_deref() {
                        Some("0") | Some("false") => false,
                        _ => true,
                    };
                    props.cant_split = Some(cant_split);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:trPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_table_borders(
    reader: &mut Reader<&[u8]>,
    end_tag: &[u8],
) -> Result<Option<TableBorders>, ParseError> {
    let mut buf = Vec::new();
    let mut borders = TableBorders::default();
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
                    b"w:insideH" => {
                        borders.inside_h = border;
                        has_any = true;
                    }
                    b"w:insideV" => {
                        borders.inside_v = border;
                        has_any = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_tag {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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

fn parse_border(start: &BytesStart) -> Option<Border> {
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

fn parse_cell_margins(reader: &mut Reader<&[u8]>) -> Result<Option<CellMargins>, ParseError> {
    let mut buf = Vec::new();
    let mut margins = CellMargins::default();
    let mut has_any = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let val = attr_value(&e, b"w:w").and_then(|v| v.parse().ok());
                match e.name().as_ref() {
                    b"w:top" => {
                        margins.top = val;
                        has_any = true;
                    }
                    b"w:bottom" => {
                        margins.bottom = val;
                        has_any = true;
                    }
                    b"w:left" => {
                        margins.left = val;
                        has_any = true;
                    }
                    b"w:right" => {
                        margins.right = val;
                        has_any = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:tblCellMar" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    if has_any {
        Ok(Some(margins))
    } else {
        Ok(None)
    }
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

fn parse_paragraph_borders(
    reader: &mut Reader<&[u8]>,
) -> Result<Option<docir_core::ir::ParagraphBorders>, ParseError> {
    let mut buf = Vec::new();
    let mut borders = docir_core::ir::ParagraphBorders::default();
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
                if e.name().as_ref() == b"w:pBdr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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

fn parse_drawing(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<Option<NodeId>, ParseError> {
    let mut buf = Vec::new();
    let mut rel_id: Option<String> = None;
    let mut chart_rel: Option<String> = None;
    let mut diagram_rel_ids: Vec<String> = Vec::new();
    let mut name: Option<String> = None;
    let mut alt_text: Option<String> = None;
    let mut shape_type = docir_core::ir::ShapeType::Picture;
    let mut transform = docir_core::ir::ShapeTransform::default();
    let mut next_pos_is_x = true;
    let mut text: Option<docir_core::ir::ShapeText> = None;
    let mut hyperlink_rel: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_bytes = e.name().as_ref().to_vec();
                let n = name_bytes.as_slice();
                if n == b"a:blip" {
                    rel_id = attr_value(&e, b"r:embed").or_else(|| attr_value(&e, b"r:link"));
                } else if n == b"wp:docPr" {
                    name = attr_value(&e, b"name");
                    alt_text = attr_value(&e, b"descr");
                } else if n == b"a:graphicData" {
                    if let Some(uri) = attr_value(&e, b"uri") {
                        if uri.contains("chart") {
                            shape_type = docir_core::ir::ShapeType::Chart;
                        } else if uri.contains("diagram") {
                            shape_type = docir_core::ir::ShapeType::Custom;
                        }
                    }
                } else if n == b"a:prstGeom" {
                    if let Some(val) = attr_value(&e, b"prst") {
                        shape_type = map_shape_type(&val);
                    }
                } else if n == b"wp:extent" || n == b"a:ext" {
                    if let Some(val) = attr_value(&e, b"cx").and_then(|v| v.parse().ok()) {
                        transform.width = val;
                    }
                    if let Some(val) = attr_value(&e, b"cy").and_then(|v| v.parse().ok()) {
                        transform.height = val;
                    }
                } else if n == b"a:off" {
                    if let Some(val) = attr_value(&e, b"x").and_then(|v| v.parse().ok()) {
                        transform.x = val;
                    }
                    if let Some(val) = attr_value(&e, b"y").and_then(|v| v.parse().ok()) {
                        transform.y = val;
                    }
                } else if n == b"wp:posOffset" {
                    if let Ok(text) = reader.read_text(e.name()) {
                        if let Ok(val) = text.parse::<i64>() {
                            if next_pos_is_x {
                                transform.x = val;
                            } else {
                                transform.y = val;
                            }
                            next_pos_is_x = !next_pos_is_x;
                        }
                    }
                } else if n == b"a:txBody" {
                    text = Some(parse_drawing_text_body(reader, "word/document.xml")?);
                } else if n.ends_with(b":chart") || n == b"c:chart" {
                    chart_rel = attr_value(&e, b"r:id");
                } else if n == b"dgm:relIds" {
                    if let Some(val) = attr_value(&e, b"r:dm") {
                        diagram_rel_ids.push(val);
                    }
                    if let Some(val) = attr_value(&e, b"r:lo") {
                        diagram_rel_ids.push(val);
                    }
                    if let Some(val) = attr_value(&e, b"r:qs") {
                        diagram_rel_ids.push(val);
                    }
                    if let Some(val) = attr_value(&e, b"r:cs") {
                        diagram_rel_ids.push(val);
                    }
                } else if n == b"a:hlinkClick" {
                    hyperlink_rel = attr_value(&e, b"r:id");
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:drawing" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    let rel_id = chart_rel
        .clone()
        .or(diagram_rel_ids.first().cloned())
        .or(rel_id);
    if let Some(rel_id) = rel_id {
        if let Some(rel) = rels.get(&rel_id) {
            let mut shape = docir_core::ir::Shape::new(shape_type);
            shape.name = name;
            shape.alt_text = alt_text;
            shape.transform = transform;
            shape.text = text;
            shape.relationship_id = Some(rel_id.clone());
            shape.media_target = Some(normalize_docx_target(&rel.target));
            let mut span = span_from_reader(reader, "word/document.xml");
            span.relationship_id = Some(rel_id.clone());
            shape.span = Some(span);
            if let Some(hrel) = hyperlink_rel.as_ref().and_then(|id| rels.get(id)) {
                shape.hyperlink = Some(hrel.target.clone());
            }
            if !diagram_rel_ids.is_empty() {
                let mut related_targets = Vec::new();
                for rel_id in diagram_rel_ids {
                    if let Some(rel) = rels.get(&rel_id) {
                        related_targets.push(normalize_docx_target(&rel.target));
                    }
                }
                shape.related_targets = related_targets;
            }
            let shape_id = shape.id;
            parser.store.insert(docir_core::ir::IRNode::Shape(shape));
            return Ok(Some(shape_id));
        }
    }
    Ok(None)
}

fn parse_drawing_text_body(
    reader: &mut Reader<&[u8]>,
    doc_path: &str,
) -> Result<docir_core::ir::ShapeText, ParseError> {
    let mut paragraphs = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"a:p" {
                    let paragraph = parse_drawing_text_paragraph(reader, doc_path)?;
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
                    file: doc_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(docir_core::ir::ShapeText { paragraphs })
}

fn parse_drawing_text_paragraph(
    reader: &mut Reader<&[u8]>,
    doc_path: &str,
) -> Result<docir_core::ir::ShapeTextParagraph, ParseError> {
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
                    let run = parse_drawing_text_run(reader, doc_path)?;
                    runs.push(run);
                }
                b"a:br" => {
                    runs.push(docir_core::ir::ShapeTextRun {
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
                    file: doc_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(docir_core::ir::ShapeTextParagraph { runs, alignment })
}

fn parse_drawing_text_run(
    reader: &mut Reader<&[u8]>,
    doc_path: &str,
) -> Result<docir_core::ir::ShapeTextRun, ParseError> {
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
                        file: doc_path.to_string(),
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
                    file: doc_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(docir_core::ir::ShapeTextRun {
        text,
        bold,
        italic,
        font_size,
        font_family,
    })
}

fn map_shape_type(value: &str) -> docir_core::ir::ShapeType {
    match value {
        "rect" => docir_core::ir::ShapeType::Rectangle,
        "roundRect" => docir_core::ir::ShapeType::RoundRect,
        "ellipse" => docir_core::ir::ShapeType::Ellipse,
        "triangle" => docir_core::ir::ShapeType::Triangle,
        "line" => docir_core::ir::ShapeType::Line,
        "straightConnector1" => docir_core::ir::ShapeType::Line,
        "bentConnector2" | "bentConnector3" | "bentConnector4" | "bentConnector5" => {
            docir_core::ir::ShapeType::Line
        }
        "rightArrow" | "leftArrow" | "upArrow" | "downArrow" | "leftRightArrow" | "upDownArrow"
        | "bentArrow" | "uTurnArrow" | "curvedRightArrow" | "curvedLeftArrow" | "curvedUpArrow"
        | "curvedDownArrow" => docir_core::ir::ShapeType::Arrow,
        _ => docir_core::ir::ShapeType::Custom,
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

fn normalize_docx_target(target: &str) -> String {
    let mut t = target;
    while t.starts_with("../") {
        t = &t[3..];
    }
    if t.starts_with("./") {
        t = &t[2..];
    }
    if t.starts_with("word/") {
        t.to_string()
    } else {
        format!("word/{}", t.trim_start_matches('/'))
    }
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
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
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

fn parse_comments_like(
    parser: &mut DocxParser,
    xml: &str,
    rels: &Relationships,
    kind: Option<NoteKind>,
) -> Result<Vec<NodeId>, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    reader.config_mut().check_end_names = false;
    let mut buf = Vec::new();
    let mut nodes = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:comment" => {
                    let comment_id = attr_value(&e, b"w:id").unwrap_or_default();
                    let mut comment = Comment::new(comment_id);
                    comment.author = attr_value(&e, b"w:author");
                    comment.initials = attr_value(&e, b"w:initials");
                    comment.parent_id = attr_value(&e, b"w:parentId");
                    comment.para_id = attr_value(&e, b"w:paraId");
                    if let Some(val) = attr_value(&e, b"w:done") {
                        let v = val.as_str();
                        comment.done = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                    }
                    comment.date = attr_value(&e, b"w:date");
                    comment.content = parse_block_until(parser, &mut reader, rels, b"w:comment")?;
                    let id = comment.id;
                    parser
                        .store
                        .insert(docir_core::ir::IRNode::Comment(comment));
                    nodes.push(id);
                }
                b"w:footnote" => {
                    if matches!(kind, Some(NoteKind::Footnote)) {
                        let note_id = attr_value(&e, b"w:id").unwrap_or_default();
                        let mut note = Footnote::new(note_id);
                        note.note_type = attr_value(&e, b"w:type");
                        note.content = parse_block_until(parser, &mut reader, rels, b"w:footnote")?;
                        let id = note.id;
                        parser.store.insert(docir_core::ir::IRNode::Footnote(note));
                        nodes.push(id);
                    }
                }
                b"w:endnote" => {
                    if matches!(kind, Some(NoteKind::Endnote)) {
                        let note_id = attr_value(&e, b"w:id").unwrap_or_default();
                        let mut note = Endnote::new(note_id);
                        note.note_type = attr_value(&e, b"w:type");
                        note.content = parse_block_until(parser, &mut reader, rels, b"w:endnote")?;
                        let id = note.id;
                        parser.store.insert(docir_core::ir::IRNode::Endnote(note));
                        nodes.push(id);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/comments.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(nodes)
}

fn parse_settings_like(xml: &str) -> Result<WordSettings, ParseError> {
    let mut settings = WordSettings::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
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
                return Err(ParseError::Xml {
                    file: "word/settings.xml".to_string(),
                    message: e.to_string(),
                });
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
                return Err(ParseError::Xml {
                    file: "word/numbering.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(abstract_id)
}

fn attr_value(start: &BytesStart, key: &[u8]) -> Option<String> {
    for attr in start.attributes().flatten() {
        if attr.key.as_ref() == key {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

fn skip_to_end(reader: &mut Reader<&[u8]>, end: &[u8]) -> Result<(), ParseError> {
    let mut depth = 0usize;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == end {
                    depth += 1;
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end {
                    if depth == 0 {
                        break;
                    }
                    depth -= 1;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
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

fn span_from_reader(reader: &Reader<&[u8]>, file_path: &str) -> SourceSpan {
    let mut span = SourceSpan::new(file_path);
    if let Ok(pos) = usize::try_from(reader.buffer_position()) {
        if let Some((line, col)) = line_col(reader.get_ref(), pos) {
            span.line = Some(line);
            span.column = Some(col);
        }
    }
    span
}
