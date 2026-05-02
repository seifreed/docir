//! DOCX document parsing (minimal but real).

use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::ooxml::shared::normalize_docx_target;
use crate::xml_utils::{attr_value, reader_from_str, scan_xml_events_with_reader, XmlScanControl};
use docir_core::ir::{
    Border, BorderStyle, CommentRangeEnd, CommentRangeStart, CommentReference, Document, Field,
    Footer, GlossaryDocument, Header, PageBorders, ParagraphProperties, Run, RunProperties,
    StyleParagraphProperties, StyleRunProperties, WebSettings, WordSettings,
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
mod inline;
mod numbering;
mod paragraph;
mod sections;
mod styles;
mod support;
mod table;
use body::{parse_block_until, parse_body_sections};
use glossary::parse_doc_part;
use inline::{parse_field, parse_hyperlink, parse_numbering, parse_run_properties};
#[cfg(test)]
use inline::{parse_run, parse_sdt, SdtMode};
#[cfg(test)]
use paragraph::parse_paragraph;
use paragraph::parse_paragraph_simple;
use sections::{apply_section_refs, SectionRef};

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

impl Default for DocxParser {
    fn default() -> Self {
        Self::new()
    }
}

impl DocxParser {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            store: IrStore::new(),
        }
    }

    pub(crate) fn store_mut(&mut self) -> &mut IrStore {
        &mut self.store
    }

    /// Public API entrypoint: into_store.
    pub fn into_store(self) -> IrStore {
        self.store
    }

    /// Public API entrypoint: parse_document.
    pub fn parse_document(
        &mut self,
        xml: &str,
        rels: &Relationships,
        header_footer_map: Option<&HashMap<String, NodeId>>,
    ) -> Result<NodeId, ParseError> {
        let mut doc = Document::new(DocumentFormat::WordProcessing);

        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();

        scan_xml_events_with_reader(
            &mut reader,
            &mut buf,
            "word/document.xml",
            |reader, event| {
                if let Event::Start(e) = event {
                    if e.name().as_ref() == b"w:body" {
                        let sections = parse_body_sections(self, reader, rels, header_footer_map)?;
                        for section in sections {
                            let section_id = section.id;
                            self.store.insert(docir_core::ir::IRNode::Section(section));
                            doc.content.push(section_id);
                        }
                    }
                }
                Ok(XmlScanControl::Continue)
            },
        )?;

        let doc_id = doc.id;
        self.store.insert(docir_core::ir::IRNode::Document(doc));
        Ok(doc_id)
    }

    /// Public API entrypoint: parse_glossary_document.
    pub fn parse_glossary_document(
        &mut self,
        xml: &str,
        rels: &Relationships,
    ) -> Result<NodeId, ParseError> {
        let mut glossary = GlossaryDocument::new();
        glossary.span = Some(SourceSpan::new("word/glossary/document.xml"));

        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();

        scan_xml_events_with_reader(
            &mut reader,
            &mut buf,
            "word/glossary/document.xml",
            |reader, event| {
                if let Event::Start(e) = event {
                    if e.name().as_ref() == b"w:docPart" {
                        let entry = parse_doc_part(self, reader, rels)?;
                        let entry_id = entry.id;
                        self.store
                            .insert(docir_core::ir::IRNode::GlossaryEntry(entry));
                        glossary.entries.push(entry_id);
                    }
                }
                Ok(XmlScanControl::Continue)
            },
        )?;

        let glossary_id = glossary.id;
        self.store
            .insert(docir_core::ir::IRNode::GlossaryDocument(glossary));
        Ok(glossary_id)
    }

    /// Public API entrypoint: parse_header_footer.
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

    /// Public API entrypoint: parse_settings.
    pub fn parse_settings(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let settings = parse_settings_like(xml)?;
        let id = settings.id;
        self.store
            .insert(docir_core::ir::IRNode::WordSettings(settings));
        Ok(id)
    }

    /// Public API entrypoint: parse_web_settings.
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
    support::parse_page_borders(reader)
}

fn bool_from_val(start: &BytesStart) -> bool {
    support::bool_from_val(start)
}

fn parse_settings_like(xml: &str) -> Result<WordSettings, ParseError> {
    support::parse_settings_like(xml)
}

fn parse_num_abstract_id(reader: &mut Reader<&[u8]>) -> Result<u32, ParseError> {
    support::parse_num_abstract_id(reader)
}

fn line_col(data: &[u8], pos: usize) -> Option<(u32, u32)> {
    support::line_col(data, pos)
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
