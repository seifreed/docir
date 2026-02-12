//! Main OOXML parser orchestrator.

pub use crate::config::{ParseMetrics, ParserConfig};
use crate::diagnostics::{push_info, push_warning};
use crate::error::ParseError;
use crate::format::FormatParser;
use crate::hwp::is_hwpx_mimetype;
use crate::input::{enforce_input_size, read_all_with_limit};
use crate::ole::is_ole_container;
use crate::ooxml::content_types::ContentTypes;
use crate::ooxml::docx::DocxParser;
use crate::ooxml::part_utils::{read_relationships_optional, read_xml_part, read_xml_part_by_rel};
use crate::ooxml::pptx::PptxParser;
use crate::ooxml::relationships::{rel_type, Relationships};
use crate::ooxml::xlsx::XlsxParser;
use crate::rtf::is_rtf_bytes;
use crate::xml_utils::local_name;
use crate::zip_handler::SecureZipReader;
use docir_core::ir::column_to_letter;
use docir_core::ir::{
    Cell, CellError, CellValue, Diagnostics, Document, ExtensionPart, ExtensionPartKind, IRNode,
    MediaAsset, SheetKind, SheetState, Worksheet,
};
use docir_core::normalize::normalize_store;
use docir_core::security::SecurityInfo;
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use std::collections::HashMap;
use std::io::{Cursor, Read, Seek, SeekFrom};
use std::path::Path;

mod coverage;
mod dispatch;
mod metadata;
mod parser_docx;
mod parser_pptx;
mod parser_xlsx;
mod post_process;
pub(crate) mod security;
mod shared_parts;
mod utils;
mod vba;

/// Parsed document result.
#[derive(Debug)]
pub struct ParsedDocument {
    /// Root document node ID.
    pub root_id: NodeId,
    /// Document format.
    pub format: DocumentFormat,
    /// IR node store.
    pub store: IrStore,
    /// Optional parse metrics.
    pub metrics: Option<ParseMetrics>,
}

struct DocxWordParts {
    styles_id: Option<NodeId>,
    styles_with_effects_id: Option<NodeId>,
    numbering_id: Option<NodeId>,
    comments: Vec<NodeId>,
    footnotes: Vec<NodeId>,
    endnotes: Vec<NodeId>,
    settings_id: Option<NodeId>,
    web_settings_id: Option<NodeId>,
    font_table_id: Option<NodeId>,
    comments_ext_id: Option<NodeId>,
    comments_id_map_id: Option<NodeId>,
    glossary_id: Option<NodeId>,
}

impl ParsedDocument {
    /// Gets the root document node.
    pub fn document(&self) -> Option<&Document> {
        if let Some(IRNode::Document(doc)) = self.store.get(self.root_id) {
            Some(doc)
        } else {
            None
        }
    }

    /// Gets the security info from the document.
    pub fn security_info(&self) -> Option<&SecurityInfo> {
        self.document().map(|d| &d.security)
    }
}

/// Main OOXML parser.
pub struct OoxmlParser {
    config: ParserConfig,
}

impl FormatParser for OoxmlParser {
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        self.parse_reader(reader)
    }
}

impl OoxmlParser {
    /// Creates a new parser with default configuration.
    pub fn new() -> Self {
        Self {
            config: ParserConfig::default(),
        }
    }

    /// Creates a new parser with custom configuration.
    pub fn with_config(config: ParserConfig) -> Self {
        Self { config }
    }

    crate::impl_parse_entrypoints!();

    /// Parses from any reader.
    pub fn parse_reader<R: Read + Seek>(
        &self,
        mut reader: R,
    ) -> Result<ParsedDocument, ParseError> {
        reader.seek(SeekFrom::Start(0))?;
        let data = read_all_with_limit(reader, self.config.max_input_size)?;
        let mut zip =
            SecureZipReader::new(Cursor::new(data.as_slice()), self.config.zip_config.clone())?;
        let mut metrics = self.init_metrics();
        let content_types = self.read_content_types(&mut zip, &mut metrics)?;
        let format = self.detect_format(&content_types)?;
        let package_rels = self.read_package_relationships(&mut zip, &mut metrics)?;
        let main_part_path = self.resolve_main_part(&package_rels)?;

        if self.config.enforce_required_parts {
            self.validate_structure(&mut zip, format, &main_part_path)?;
        }

        let mut parsed = self.parse_main(
            &mut zip,
            data.as_slice(),
            format,
            &main_part_path,
            &content_types,
            &mut metrics,
        )?;
        if let Some(m) = metrics {
            parsed.metrics = Some(m);
        }
        Ok(parsed)
    }

    fn init_metrics(&self) -> Option<ParseMetrics> {
        if self.config.enable_metrics {
            Some(ParseMetrics::default())
        } else {
            None
        }
    }

    fn read_content_types<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        metrics: &mut Option<ParseMetrics>,
    ) -> Result<ContentTypes, ParseError> {
        let start = std::time::Instant::now();
        let content_types_xml = zip.read_file_string("[Content_Types].xml")?;
        let content_types = ContentTypes::parse(&content_types_xml)?;
        if let Some(m) = metrics.as_mut() {
            m.content_types_ms = start.elapsed().as_millis();
        }
        Ok(content_types)
    }

    fn detect_format(&self, content_types: &ContentTypes) -> Result<DocumentFormat, ParseError> {
        content_types
            .detect_format()
            .ok_or_else(|| ParseError::UnsupportedFormat("Unknown OOXML format".to_string()))
    }

    fn read_package_relationships<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        metrics: &mut Option<ParseMetrics>,
    ) -> Result<Relationships, ParseError> {
        let start = std::time::Instant::now();
        let rels_xml = zip.read_file_string("_rels/.rels")?;
        let rels = Relationships::parse(&rels_xml)?;
        if let Some(m) = metrics.as_mut() {
            m.relationships_ms = start.elapsed().as_millis();
        }
        Ok(rels)
    }

    fn resolve_main_part(&self, package_rels: &Relationships) -> Result<String, ParseError> {
        let main_rel = package_rels
            .get_first_by_type(rel_type::OFFICE_DOCUMENT)
            .ok_or_else(|| ParseError::MissingPart("Office document relationship".to_string()))?;
        Ok(Relationships::resolve_target("", &main_rel.target))
    }

    fn parse_main<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        data: &[u8],
        format: DocumentFormat,
        main_part_path: &str,
        content_types: &ContentTypes,
        metrics: &mut Option<ParseMetrics>,
    ) -> Result<ParsedDocument, ParseError> {
        let start = std::time::Instant::now();
        let parsed = match format {
            DocumentFormat::WordProcessing => {
                self.parse_docx(zip, main_part_path, content_types, metrics)
            }
            DocumentFormat::Spreadsheet => {
                if main_part_path.ends_with(".bin") {
                    self.parse_xlsb(zip, data, content_types, metrics)
                } else {
                    self.parse_xlsx(zip, main_part_path, content_types, metrics)
                }
            }
            DocumentFormat::Presentation => {
                self.parse_pptx(zip, main_part_path, content_types, metrics)
            }
            _ => {
                return Err(ParseError::UnsupportedFormat(
                    "Non-OOXML format requested".to_string(),
                ))
            }
        }?;
        if let Some(m) = metrics.as_mut() {
            m.main_parse_ms = start.elapsed().as_millis();
        }
        Ok(parsed)
    }

    fn validate_structure<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        format: DocumentFormat,
        main_part_path: &str,
    ) -> Result<(), ParseError> {
        if !zip.contains(main_part_path) {
            return Err(ParseError::MissingPart(main_part_path.to_string()));
        }

        match format {
            DocumentFormat::WordProcessing => {
                if !main_part_path.starts_with("word/") {
                    return Err(ParseError::InvalidFormat(format!(
                        "Expected Word document under word/, got {main_part_path}"
                    )));
                }
            }
            DocumentFormat::Spreadsheet => {
                if !main_part_path.starts_with("xl/") {
                    return Err(ParseError::InvalidFormat(format!(
                        "Expected Excel workbook under xl/, got {main_part_path}"
                    )));
                }
            }
            DocumentFormat::Presentation => {
                if !main_part_path.starts_with("ppt/") {
                    return Err(ParseError::InvalidFormat(format!(
                        "Expected PPTX presentation under ppt/, got {main_part_path}"
                    )));
                }
            }
            _ => {
                return Err(ParseError::UnsupportedFormat(
                    "Non-OOXML format requested".to_string(),
                ))
            }
        }
        Ok(())
    }

    fn parse_docx_word_parts<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> DocxWordParts {
        let styles_id = self.parse_docx_part_by_rel_with_span(
            zip,
            main_part_path,
            doc_rels,
            rel_type::STYLES,
            parser,
            |parser, _part_path, xml| parser.parse_styles(xml).ok(),
            |store, id, part_path| {
                if let Some(IRNode::StyleSet(set)) = store.get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
            },
        );

        let styles_with_effects_id = self.parse_docx_part_by_path_with_span(
            zip,
            "word/stylesWithEffects.xml",
            parser,
            |parser, _part_path, xml| parser.parse_styles_with_effects(xml).ok(),
            |store, id, part_path| {
                if let Some(IRNode::StyleSet(set)) = store.get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
            },
        );

        let numbering_id = self.parse_docx_part_by_rel_with_span(
            zip,
            main_part_path,
            doc_rels,
            rel_type::NUMBERING,
            parser,
            |parser, _part_path, xml| parser.parse_numbering(xml).ok(),
            |store, id, part_path| {
                if let Some(IRNode::NumberingSet(set)) = store.get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
            },
        );

        let comments = self.parse_docx_comments(zip, main_part_path, doc_rels, parser);

        let footnotes = self.parse_docx_notes(
            zip,
            main_part_path,
            doc_rels,
            parser,
            rel_type::FOOTNOTES,
            crate::ooxml::docx::document::NoteKind::Footnote,
        );

        let endnotes = self.parse_docx_notes(
            zip,
            main_part_path,
            doc_rels,
            parser,
            rel_type::ENDNOTES,
            crate::ooxml::docx::document::NoteKind::Endnote,
        );

        let settings_id = self.parse_docx_part_by_rel_with_span(
            zip,
            main_part_path,
            doc_rels,
            rel_type::SETTINGS,
            parser,
            |parser, _part_path, xml| parser.parse_settings(xml).ok(),
            |store, id, part_path| {
                if let Some(IRNode::WordSettings(settings)) = store.get_mut(id) {
                    settings.span = Some(SourceSpan::new(part_path));
                }
            },
        );

        let web_settings_id = self.parse_docx_part_by_rel_with_span(
            zip,
            main_part_path,
            doc_rels,
            rel_type::WEB_SETTINGS,
            parser,
            |parser, _part_path, xml| parser.parse_web_settings(xml).ok(),
            |store, id, part_path| {
                if let Some(IRNode::WebSettings(settings)) = store.get_mut(id) {
                    settings.span = Some(SourceSpan::new(part_path));
                }
            },
        );

        let font_table_id = self.parse_docx_font_table(zip, main_part_path, doc_rels, parser);

        let comments_ext_id = self.parse_docx_part_by_path_with_span(
            zip,
            "word/commentsExtended.xml",
            parser,
            |parser, _part_path, xml| parser.parse_comments_extended(xml).ok(),
            |store, id, part_path| {
                if let Some(IRNode::CommentExtensionSet(set)) = store.get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
            },
        );

        let comments_id_map_id = self.parse_docx_part_by_path_with_span(
            zip,
            "word/commentsIds.xml",
            parser,
            |parser, _part_path, xml| parser.parse_comments_ids(xml).ok(),
            |store, id, part_path| {
                if let Some(IRNode::CommentIdMap(map)) = store.get_mut(id) {
                    map.span = Some(SourceSpan::new(part_path));
                }
            },
        );

        let glossary_id =
            self.parse_docx_part_by_path(zip, "word/glossary/document.xml", |_, xml| {
                parser.parse_glossary_document(xml, doc_rels).ok()
            });

        DocxWordParts {
            styles_id,
            styles_with_effects_id,
            numbering_id,
            comments,
            footnotes,
            endnotes,
            settings_id,
            web_settings_id,
            font_table_id,
            comments_ext_id,
            comments_id_map_id,
            glossary_id,
        }
    }

    fn parse_docx_headers_footers<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> Result<HashMap<String, NodeId>, ParseError> {
        let mut map = HashMap::new();

        for rel in doc_rels.get_by_type(rel_type::HEADER) {
            let part_path = Relationships::resolve_target(main_part_path, &rel.target);
            let rels = read_relationships_optional(zip, &part_path);
            let xml = zip.read_file_string(&part_path)?;
            let node_id = parser.parse_header_footer(
                &xml,
                &part_path,
                crate::ooxml::docx::document::HeaderFooterKind::Header,
                &rels,
            )?;
            map.insert(rel.id.clone(), node_id);
        }

        for rel in doc_rels.get_by_type(rel_type::FOOTER) {
            let part_path = Relationships::resolve_target(main_part_path, &rel.target);
            let rels = read_relationships_optional(zip, &part_path);
            let xml = zip.read_file_string(&part_path)?;
            let node_id = parser.parse_header_footer(
                &xml,
                &part_path,
                crate::ooxml::docx::document::HeaderFooterKind::Footer,
                &rels,
            )?;
            map.insert(rel.id.clone(), node_id);
        }

        Ok(map)
    }

    fn parse_docx_comments<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> Vec<NodeId> {
        let comments = doc_rels
            .get_first_by_type(rel_type::COMMENTS)
            .and_then(|rel| {
                let part_path = Relationships::resolve_target(main_part_path, &rel.target);
                let rels = read_relationships_optional(zip, &part_path);
                zip.read_file_string(&part_path).ok().and_then(|xml| {
                    let ids = parser.parse_comments(&xml, &rels).ok()?;
                    for id in &ids {
                        if let Some(IRNode::Comment(comment)) = parser.store_mut().get_mut(*id) {
                            comment.span = Some(SourceSpan::new(&part_path));
                        }
                    }
                    Some(ids)
                })
            })
            .unwrap_or_default();

        if !comments.is_empty() || !zip.contains("word/comments.xml") {
            return comments;
        }

        let rels = read_relationships_optional(zip, "word/comments.xml");
        if let Ok(xml) = zip.read_file_string("word/comments.xml") {
            if let Ok(ids) = parser.parse_comments(&xml, &rels) {
                for id in &ids {
                    if let Some(IRNode::Comment(comment)) = parser.store_mut().get_mut(*id) {
                        comment.span = Some(SourceSpan::new("word/comments.xml"));
                    }
                }
                return ids;
            }
        }

        comments
    }

    fn parse_docx_notes<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
        rel_type: &str,
        kind: crate::ooxml::docx::document::NoteKind,
    ) -> Vec<NodeId> {
        doc_rels
            .get_first_by_type(rel_type)
            .and_then(|rel| {
                let part_path = Relationships::resolve_target(main_part_path, &rel.target);
                let rels = read_relationships_optional(zip, &part_path);
                zip.read_file_string(&part_path).ok().and_then(|xml| {
                    let ids = parser.parse_notes(&xml, kind, &rels).ok()?;
                    for id in &ids {
                        match kind {
                            crate::ooxml::docx::document::NoteKind::Footnote => {
                                if let Some(IRNode::Footnote(note)) =
                                    parser.store_mut().get_mut(*id)
                                {
                                    note.span = Some(SourceSpan::new(&part_path));
                                }
                            }
                            crate::ooxml::docx::document::NoteKind::Endnote => {
                                if let Some(IRNode::Endnote(note)) = parser.store_mut().get_mut(*id)
                                {
                                    note.span = Some(SourceSpan::new(&part_path));
                                }
                            }
                        }
                    }
                    Some(ids)
                })
            })
            .unwrap_or_default()
    }

    fn parse_docx_font_table<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> Option<NodeId> {
        let mut font_table_id = self.parse_docx_part_by_rel_with_span(
            zip,
            main_part_path,
            doc_rels,
            rel_type::FONT_TABLE,
            parser,
            |parser, _part_path, xml| parser.parse_font_table(xml).ok(),
            |store, id, part_path| {
                if let Some(IRNode::FontTable(table)) = store.get_mut(id) {
                    table.span = Some(SourceSpan::new(part_path));
                }
            },
        );

        if font_table_id.is_none() && zip.contains("word/fontTable.xml") {
            if let Ok(xml) = zip.read_file_string("word/fontTable.xml") {
                if let Ok(id) = parser.parse_font_table(&xml) {
                    if let Some(IRNode::FontTable(table)) = parser.store_mut().get_mut(id) {
                        table.span = Some(SourceSpan::new("word/fontTable.xml"));
                    }
                    font_table_id = Some(id);
                }
            }
        }

        font_table_id
    }

    fn parse_docx_part_by_rel<R, F>(
        &self,
        zip: &mut SecureZipReader<R>,
        main_part_path: &str,
        doc_rels: &Relationships,
        rel_type: &str,
        mut parse: F,
    ) -> Option<NodeId>
    where
        R: Read + Seek,
        F: FnMut(&str, &str) -> Option<NodeId>,
    {
        let (part_path, xml) =
            read_xml_part_by_rel(zip, main_part_path, doc_rels, rel_type).ok()??;
        parse(&part_path, &xml)
    }

    fn parse_docx_part_by_rel_with_span<R, F, S>(
        &self,
        zip: &mut SecureZipReader<R>,
        main_part_path: &str,
        doc_rels: &Relationships,
        rel_type: &str,
        parser: &mut DocxParser,
        parse: F,
        set_span: S,
    ) -> Option<NodeId>
    where
        R: Read + Seek,
        F: FnOnce(&mut DocxParser, &str, &str) -> Option<NodeId>,
        S: FnOnce(&mut IrStore, NodeId, &str),
    {
        let (part_path, xml) =
            read_xml_part_by_rel(zip, main_part_path, doc_rels, rel_type).ok()??;
        let id = parse(parser, &part_path, &xml)?;
        set_span(parser.store_mut(), id, &part_path);
        Some(id)
    }

    fn parse_docx_part_by_path<R, F>(
        &self,
        zip: &mut SecureZipReader<R>,
        part_path: &str,
        mut parse: F,
    ) -> Option<NodeId>
    where
        R: Read + Seek,
        F: FnMut(&str, &str) -> Option<NodeId>,
    {
        let xml = read_xml_part(zip, part_path).ok()??;
        parse(part_path, &xml)
    }

    fn parse_docx_part_by_path_with_span<R, F, S>(
        &self,
        zip: &mut SecureZipReader<R>,
        part_path: &str,
        parser: &mut DocxParser,
        parse: F,
        set_span: S,
    ) -> Option<NodeId>
    where
        R: Read + Seek,
        F: FnOnce(&mut DocxParser, &str, &str) -> Option<NodeId>,
        S: FnOnce(&mut IrStore, NodeId, &str),
    {
        let xml = read_xml_part(zip, part_path).ok()??;
        let id = parse(parser, part_path, &xml)?;
        set_span(parser.store_mut(), id, part_path);
        Some(id)
    }
}

/// Unified parser that dispatches OOXML vs ODF based on package signature.
pub struct DocumentParser {
    config: ParserConfig,
}

impl DocumentParser {
    /// Creates a new parser with default configuration.
    pub fn new() -> Self {
        Self {
            config: ParserConfig::default(),
        }
    }

    /// Creates a new parser with custom configuration.
    pub fn with_config(config: ParserConfig) -> Self {
        Self { config }
    }

    crate::impl_parse_entrypoints!();

    /// Parses from any reader.
    pub fn parse_reader<R: Read + Seek>(
        &self,
        mut reader: R,
    ) -> Result<ParsedDocument, ParseError> {
        enforce_input_size(&mut reader, self.config.max_input_size)?;
        let detected = self.detect_format(&mut reader)?;
        self.validate_detected(&detected, &mut reader)?;
        let parsed = self.parse_detected(detected, reader)?;
        self.normalize_parsed(parsed)
    }

    fn detect_format<R: Read + Seek>(
        &self,
        reader: &mut R,
    ) -> Result<dispatch::DetectedFormat, ParseError> {
        let mut probe = [0u8; 16];
        let read = reader.read(&mut probe)?;
        reader.seek(SeekFrom::Start(0))?;
        let head = &probe[..read];

        if is_rtf_bytes(head) {
            return Ok(dispatch::DetectedFormat::Rtf);
        }

        if is_ole_container(head) {
            return Ok(dispatch::DetectedFormat::Hwp);
        }

        if !is_zip_container(head) {
            return Err(ParseError::UnsupportedFormat(
                "Unknown package format (not OLE/CFB or ZIP)".to_string(),
            ));
        }

        let mut inspector = SecureZipReader::new(&mut *reader, self.config.zip_config.clone())?;
        let is_ooxml = inspector.contains("[Content_Types].xml");
        let is_odf = inspector.contains("mimetype");
        let is_hwpx = if is_odf {
            inspector
                .read_file_string("mimetype")
                .map(|m| is_hwpx_mimetype(&m))
                .unwrap_or(false)
        } else {
            false
        };
        drop(inspector);
        reader.seek(SeekFrom::Start(0))?;

        if is_ooxml {
            return Ok(dispatch::DetectedFormat::Ooxml);
        }
        if is_hwpx {
            return Ok(dispatch::DetectedFormat::Hwpx);
        }
        if is_odf {
            return Ok(dispatch::DetectedFormat::Odf);
        }

        Err(ParseError::UnsupportedFormat(
            "Unknown package format (missing [Content_Types].xml and mimetype)".to_string(),
        ))
    }

    fn validate_detected<R: Read + Seek>(
        &self,
        _detected: &dispatch::DetectedFormat,
        _reader: &mut R,
    ) -> Result<(), ParseError> {
        Ok(())
    }

    fn parse_detected<R: Read + Seek>(
        &self,
        detected: dispatch::DetectedFormat,
        reader: R,
    ) -> Result<ParsedDocument, ParseError> {
        dispatch::build_parser(detected, self.config.clone()).parse_reader(reader)
    }

    fn normalize_parsed(&self, parsed: ParsedDocument) -> Result<ParsedDocument, ParseError> {
        Ok(parsed)
    }

    /// Parses from a file and returns parsed document with raw bytes.
    pub fn parse_file_with_bytes<P: AsRef<Path>>(
        &self,
        path: P,
    ) -> Result<(ParsedDocument, Vec<u8>), ParseError> {
        let reader = crate::input::open_reader(path)?;
        let data = read_all_with_limit(reader, self.config.max_input_size)?;
        let parsed = self.parse_bytes(&data)?;
        Ok((parsed, data))
    }

    /// Parses from a reader and returns parsed document with raw bytes.
    pub fn parse_reader_with_bytes<R: Read + Seek>(
        &self,
        reader: R,
    ) -> Result<(ParsedDocument, Vec<u8>), ParseError> {
        let data = read_all_with_limit(reader, self.config.max_input_size)?;
        let parsed = self.parse_bytes(&data)?;
        Ok((parsed, data))
    }
}

fn is_zip_container(data: &[u8]) -> bool {
    data.len() >= 4 && data[0] == b'P' && data[1] == b'K'
}

fn parse_activex_xml(xml: &str, _path: &str) -> Option<docir_core::security::ActiveXControl> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut control = docir_core::security::ActiveXControl::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    match key {
                        b"name" => control.name = Some(val.clone()),
                        b"clsid" | b"classid" => control.clsid = Some(val.clone()),
                        b"progid" => control.prog_id = Some(val.clone()),
                        _ => {
                            let k = String::from_utf8_lossy(key).to_string();
                            control.properties.push((k, val));
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

    if control.name.is_some() || control.clsid.is_some() || control.prog_id.is_some() {
        Some(control)
    } else {
        None
    }
}

fn map_calamine_error(err: calamine::CellErrorType) -> CellError {
    use calamine::CellErrorType::*;
    match err {
        Null => CellError::Null,
        Div0 => CellError::DivZero,
        Value => CellError::Value,
        Ref => CellError::Ref,
        Name => CellError::Name,
        Num => CellError::Num,
        NA => CellError::NA,
        GettingData => CellError::GettingData,
    }
}

fn parse_smartart_part(xml: &str, path: &str) -> Result<docir_core::ir::SmartArtPart, ParseError> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut root_element: Option<String> = None;
    let mut point_count: u32 = 0;
    let mut connection_count: u32 = 0;
    let mut rel_ids: Vec<String> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if root_element.is_none() {
                    root_element = Some(String::from_utf8_lossy(e.name().as_ref()).to_string());
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

    Ok(docir_core::ir::SmartArtPart {
        id: NodeId::new(),
        kind: "diagram".to_string(),
        path: path.to_string(),
        root_element,
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

fn parse_chart_data(
    xml: &str,
    chart_path: &str,
    store: &mut IrStore,
) -> Result<NodeId, ParseError> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut chart = docir_core::ir::ChartData::new();
    chart.span = Some(SourceSpan::new(chart_path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_title = false;
    let mut in_series = false;
    let mut section: Option<&[u8]> = None;
    let mut current_series: Option<docir_core::ir::ChartSeries> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name_bytes = e.name().as_ref().to_vec();
                let name = local_name(&name_bytes);
                if name.ends_with(b"Chart") || name.ends_with(b"chart") {
                    chart.chart_type = Some(String::from_utf8_lossy(name).to_string());
                }
                if name == b"ser" {
                    in_series = true;
                    current_series = Some(docir_core::ir::ChartSeries::new());
                }
                if !in_series && (name == b"title") {
                    in_title = true;
                }
                if in_series {
                    if name == b"tx" {
                        section = Some(b"tx");
                    } else if name == b"cat" {
                        section = Some(b"cat");
                    } else if name == b"val" {
                        section = Some(b"val");
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name_bytes = e.name().as_ref().to_vec();
                let name = local_name(&name_bytes);
                if name == b"title" {
                    in_title = false;
                }
                if name == b"ser" {
                    in_series = false;
                    section = None;
                    if let Some(series) = current_series.take() {
                        if let Some(name) = &series.name {
                            chart.series.push(name.clone());
                        }
                        chart.series_data.push(series);
                    }
                }
                if name == b"tx" || name == b"cat" || name == b"val" {
                    section = None;
                }
            }
            Ok(Event::Text(e)) => {
                let text = e.unescape().unwrap_or_default().to_string();
                if in_title && chart.title.is_none() {
                    chart.title = Some(text);
                } else if in_series {
                    let trimmed = text.trim();
                    if trimmed.is_empty() {
                        // skip
                    } else if let Some(series) = current_series.as_mut() {
                        match section {
                            Some(b"tx") => {
                                if series.name.is_none() {
                                    series.name = Some(trimmed.to_string());
                                }
                            }
                            Some(b"cat") => {
                                series.categories.push(trimmed.to_string());
                            }
                            Some(b"val") => {
                                series.values.push(trimmed.to_string());
                            }
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: chart_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    let id = chart.id;
    store.insert(IRNode::ChartData(chart));
    Ok(id)
}

/// Convert a .rels path to its source part path when possible.

impl Default for OoxmlParser {
    fn default() -> Self {
        Self::new()
    }
}

/// Helper function to encode bytes as hex.
mod hex {
    pub fn encode(data: impl AsRef<[u8]>) -> String {
        data.as_ref().iter().map(|b| format!("{:02x}", b)).collect()
    }
}

#[cfg(test)]
mod tests;
