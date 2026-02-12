//! Main OOXML parser orchestrator.

use crate::error::ParseError;
use crate::format::FormatParser;
use crate::hwp::{is_hwpx_mimetype, HwpParser, HwpxParser};
use crate::input::{enforce_input_size, read_all_with_limit};
use crate::odf::OdfParser;
use crate::ole::is_ole_container;
use crate::ooxml::content_types::ContentTypes;
use crate::ooxml::docx::DocxParser;
use crate::ooxml::part_utils::read_relationships_optional;
use crate::ooxml::pptx::PptxParser;
use crate::ooxml::relationships::{rel_type, Relationships};
use crate::ooxml::xlsx::XlsxParser;
use crate::rtf::{is_rtf_bytes, RtfParser};
use crate::xml_utils::local_name;
use crate::zip_handler::{SecureZipReader, ZipConfig};
use docir_core::ir::column_to_letter;
use docir_core::ir::{
    Cell, CellError, CellValue, DiagnosticEntry, DiagnosticSeverity, Diagnostics, Document,
    ExtensionPart, ExtensionPartKind, IRNode, MediaAsset, SheetKind, SheetState, Worksheet,
};
use docir_core::normalize::normalize_store;
use docir_core::security::SecurityInfo;
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use std::collections::HashMap;
use std::io::{Cursor, Read, Seek, SeekFrom};

mod coverage;
mod metadata;
mod parser_docx;
mod parser_pptx;
mod parser_xlsx;
mod security;
mod shared_parts;

/// ODF parsing configuration.
#[derive(Debug, Clone)]
pub struct OdfConfig {
    /// ODF fast-mode threshold (content.xml size in bytes).
    pub fast_threshold_bytes: u64,
    /// Force ODF fast mode regardless of size.
    pub force_fast: bool,
    /// ODF fast-mode sample rows (0 = no sampling).
    pub fast_sample_rows: u32,
    /// ODF fast-mode sample columns (0 = no sampling).
    pub fast_sample_cols: u32,
    /// ODF maximum cells (None = unlimited).
    pub max_cells: Option<u64>,
    /// ODF maximum rows (None = unlimited).
    pub max_rows: Option<u64>,
    /// ODF maximum paragraphs (None = unlimited).
    pub max_paragraphs: Option<u64>,
    /// ODF maximum content.xml bytes (None = unlimited).
    pub max_bytes: Option<u64>,
    /// Enable parallel parsing of ODF sheets when possible.
    pub parallel_sheets: bool,
    /// Max threads for parallel ODF sheet parsing (None = auto).
    pub parallel_max_threads: Option<usize>,
    /// Password for decrypting encrypted ODF parts.
    pub password: Option<String>,
}

impl Default for OdfConfig {
    fn default() -> Self {
        Self {
            fast_threshold_bytes: 50 * 1024 * 1024,
            force_fast: false,
            fast_sample_rows: 0,
            fast_sample_cols: 0,
            max_cells: None,
            max_rows: None,
            max_paragraphs: None,
            max_bytes: None,
            parallel_sheets: false,
            parallel_max_threads: None,
            password: None,
        }
    }
}

/// RTF parsing configuration.
#[derive(Debug, Clone)]
pub struct RtfConfig {
    /// RTF max group depth (0 = unlimited).
    pub max_group_depth: usize,
    /// RTF max object hex length (0 = unlimited).
    pub max_object_hex_len: usize,
}

impl Default for RtfConfig {
    fn default() -> Self {
        Self {
            max_group_depth: 256,
            max_object_hex_len: 64 * 1024 * 1024,
        }
    }
}

/// HWP parsing configuration.
#[derive(Debug, Clone)]
pub struct HwpConfig {
    /// Force parse encrypted HWP streams.
    pub force_parse_encrypted: bool,
    /// Password for decrypting encrypted HWP streams.
    pub password: Option<String>,
    /// Dump HWP stream metadata (hash, size, compression).
    pub dump_streams: bool,
}

impl Default for HwpConfig {
    fn default() -> Self {
        Self {
            force_parse_encrypted: false,
            password: None,
            dump_streams: false,
        }
    }
}

/// Parser configuration.
#[derive(Debug, Clone)]
pub struct ParserConfig {
    /// ZIP extraction limits.
    pub zip_config: ZipConfig,
    /// Maximum XML recursion depth.
    pub max_xml_depth: usize,
    /// Extract VBA macro source code.
    pub extract_macro_source: bool,
    /// Compute hashes for OLE objects.
    pub compute_hashes: bool,
    /// Enforce minimal required parts per format.
    pub enforce_required_parts: bool,
    /// Collect parse timing metrics.
    pub enable_metrics: bool,
    /// Maximum allowed input size for top-level parser entrypoints.
    pub max_input_size: u64,
    /// ODF-specific configuration.
    pub odf: OdfConfig,
    /// RTF-specific configuration.
    pub rtf: RtfConfig,
    /// HWP-specific configuration.
    pub hwp: HwpConfig,
}

impl Default for ParserConfig {
    fn default() -> Self {
        Self {
            zip_config: ZipConfig::default(),
            max_xml_depth: 100,
            extract_macro_source: false,
            compute_hashes: true,
            enforce_required_parts: true,
            enable_metrics: false,
            max_input_size: 512 * 1024 * 1024,
            odf: OdfConfig::default(),
            rtf: RtfConfig::default(),
            hwp: HwpConfig::default(),
        }
    }
}

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

/// Basic parse metrics (timings in milliseconds).
#[derive(Debug, Clone, Default)]
pub struct ParseMetrics {
    pub content_types_ms: u128,
    pub relationships_ms: u128,
    pub main_parse_ms: u128,
    pub shared_parts_ms: u128,
    pub security_scan_ms: u128,
    pub extension_parts_ms: u128,
    pub normalization_ms: u128,
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
        let mut metrics = if self.config.enable_metrics {
            Some(ParseMetrics::default())
        } else {
            None
        };

        // Parse [Content_Types].xml
        let start = std::time::Instant::now();
        let content_types_xml = zip.read_file_string("[Content_Types].xml")?;
        let content_types = ContentTypes::parse(&content_types_xml)?;
        if let Some(m) = metrics.as_mut() {
            m.content_types_ms = start.elapsed().as_millis();
        }

        // Detect format
        let format = content_types
            .detect_format()
            .ok_or_else(|| ParseError::UnsupportedFormat("Unknown OOXML format".to_string()))?;

        // Parse package relationships
        let start = std::time::Instant::now();
        let rels_xml = zip.read_file_string("_rels/.rels")?;
        let package_rels = Relationships::parse(&rels_xml)?;
        if let Some(m) = metrics.as_mut() {
            m.relationships_ms = start.elapsed().as_millis();
        }

        // Find main document part
        let main_rel = package_rels
            .get_first_by_type(rel_type::OFFICE_DOCUMENT)
            .ok_or_else(|| ParseError::MissingPart("Office document relationship".to_string()))?;

        let main_part_path = Relationships::resolve_target("", &main_rel.target);

        if self.config.enforce_required_parts {
            self.validate_structure(&mut zip, format, &main_part_path)?;
        }

        // Parse based on format
        let start = std::time::Instant::now();
        let mut parsed = match format {
            DocumentFormat::WordProcessing => {
                self.parse_docx(&mut zip, &main_part_path, &content_types, &mut metrics)
            }
            DocumentFormat::Spreadsheet => {
                if main_part_path.ends_with(".bin") {
                    self.parse_xlsb(&mut zip, data.as_slice(), &content_types, &mut metrics)
                } else {
                    self.parse_xlsx(&mut zip, &main_part_path, &content_types, &mut metrics)
                }
            }
            DocumentFormat::Presentation => {
                self.parse_pptx(&mut zip, &main_part_path, &content_types, &mut metrics)
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
        if let Some(m) = metrics {
            parsed.metrics = Some(m);
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
    ) -> (
        Option<NodeId>,
        Option<NodeId>,
        Option<NodeId>,
        Vec<NodeId>,
        Vec<NodeId>,
        Vec<NodeId>,
        Option<NodeId>,
        Option<NodeId>,
        Option<NodeId>,
        Option<NodeId>,
        Option<NodeId>,
        Option<NodeId>,
    ) {
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
        let comments = if comments.is_empty() && zip.contains("word/comments.xml") {
            let rels = read_relationships_optional(zip, "word/comments.xml");
            if let Ok(xml) = zip.read_file_string("word/comments.xml") {
                if let Ok(ids) = parser.parse_comments(&xml, &rels) {
                    for id in &ids {
                        if let Some(IRNode::Comment(comment)) = parser.store_mut().get_mut(*id) {
                            comment.span = Some(SourceSpan::new("word/comments.xml"));
                        }
                    }
                    ids
                } else {
                    comments
                }
            } else {
                comments
            }
        } else {
            comments
        };

        let footnotes = doc_rels
            .get_first_by_type(rel_type::FOOTNOTES)
            .and_then(|rel| {
                let part_path = Relationships::resolve_target(main_part_path, &rel.target);
                let rels = read_relationships_optional(zip, &part_path);
                zip.read_file_string(&part_path).ok().and_then(|xml| {
                    let ids = parser
                        .parse_notes(
                            &xml,
                            crate::ooxml::docx::document::NoteKind::Footnote,
                            &rels,
                        )
                        .ok()?;
                    for id in &ids {
                        if let Some(IRNode::Footnote(note)) = parser.store_mut().get_mut(*id) {
                            note.span = Some(SourceSpan::new(&part_path));
                        }
                    }
                    Some(ids)
                })
            })
            .unwrap_or_default();

        let endnotes = doc_rels
            .get_first_by_type(rel_type::ENDNOTES)
            .and_then(|rel| {
                let part_path = Relationships::resolve_target(main_part_path, &rel.target);
                let rels = read_relationships_optional(zip, &part_path);
                zip.read_file_string(&part_path).ok().and_then(|xml| {
                    let ids = parser
                        .parse_notes(&xml, crate::ooxml::docx::document::NoteKind::Endnote, &rels)
                        .ok()?;
                    for id in &ids {
                        if let Some(IRNode::Endnote(note)) = parser.store_mut().get_mut(*id) {
                            note.span = Some(SourceSpan::new(&part_path));
                        }
                    }
                    Some(ids)
                })
            })
            .unwrap_or_default();

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

        (
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
        )
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
        let rel = doc_rels.get_first_by_type(rel_type)?;
        let part_path = Relationships::resolve_target(main_part_path, &rel.target);
        let xml = zip.read_file_string(&part_path).ok()?;
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
        let rel = doc_rels.get_first_by_type(rel_type)?;
        let part_path = Relationships::resolve_target(main_part_path, &rel.target);
        let xml = zip.read_file_string(&part_path).ok()?;
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
        if !zip.contains(part_path) {
            return None;
        }
        let xml = zip.read_file_string(part_path).ok()?;
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
        if !zip.contains(part_path) {
            return None;
        }
        let xml = zip.read_file_string(part_path).ok()?;
        let id = parse(parser, part_path, &xml)?;
        set_span(parser.store_mut(), id, part_path);
        Some(id)
    }

    fn add_extension_parts_and_diagnostics<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        content_types: &ContentTypes,
        store: &mut IrStore,
        root_id: NodeId,
    ) -> Result<(), ParseError> {
        let mut seen_paths = coverage::collect_seen_paths(store);
        seen_paths.insert("[Content_Types].xml".to_string());
        if let Some(IRNode::Document(doc)) = store.get(root_id) {
            if doc.metadata.is_some() {
                if zip.contains("docProps/core.xml") {
                    seen_paths.insert("docProps/core.xml".to_string());
                }
                if zip.contains("docProps/app.xml") {
                    seen_paths.insert("docProps/app.xml".to_string());
                }
                if zip.contains("docProps/custom.xml") {
                    seen_paths.insert("docProps/custom.xml".to_string());
                }
            }
            match doc.format {
                DocumentFormat::WordProcessing => {
                    if zip.contains("word/comments.xml") {
                        seen_paths.insert("word/comments.xml".to_string());
                    }
                }
                DocumentFormat::Presentation => {
                    if zip.contains("ppt/commentAuthors.xml") {
                        seen_paths.insert("ppt/commentAuthors.xml".to_string());
                    }
                }
                _ => {}
            }
        }

        let mut extension_ids = Vec::new();
        let mut diagnostics = Diagnostics::new();
        diagnostics.span = Some(SourceSpan::new("package"));
        let mut unparsed_count: usize = 0;

        let all_paths: Vec<String> = zip
            .file_names()
            .filter(|p| !p.ends_with('/') && !p.starts_with("[trash]/"))
            .map(|s| s.to_string())
            .collect();
        for path in all_paths {
            if seen_paths.contains(&path) {
                continue;
            }
            let data = zip.read_file(&path)?;
            let mut part = ExtensionPart::new(&path, data.len() as u64, ExtensionPartKind::Unknown);
            part.content_type = content_types.get_content_type(&path).map(|s| s.to_string());
            part.span = Some(SourceSpan::new(&path));
            let part_id = part.id;
            store.insert(IRNode::ExtensionPart(part));
            extension_ids.push(part_id);
            unparsed_count += 1;

            diagnostics.entries.push(DiagnosticEntry {
                severity: DiagnosticSeverity::Warning,
                code: "UNPARSED_PART".to_string(),
                message: format!("No parser registered for part: {}", path),
                path: Some(path),
            });
        }

        let format = match store.get(root_id) {
            Some(IRNode::Document(doc)) => doc.format,
            _ => DocumentFormat::WordProcessing,
        };
        let coverage_entries = coverage::build_coverage_diagnostics(
            format,
            content_types,
            zip.file_names().map(|s| s.to_string()).collect(),
            &seen_paths,
        );
        diagnostics.entries.extend(coverage_entries);
        diagnostics.entries.push(DiagnosticEntry {
            severity: DiagnosticSeverity::Info,
            code: "UNPARSED_SUMMARY".to_string(),
            message: format!("unparsed parts: {}", unparsed_count),
            path: None,
        });

        if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
            doc.shared_parts.extend(extension_ids);
        }

        if !diagnostics.entries.is_empty() {
            let diag_id = diagnostics.id;
            store.insert(IRNode::Diagnostics(diagnostics));
            if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
                doc.diagnostics.push(diag_id);
            }
        }

        Ok(())
    }

    /// Scan for security-relevant content.
    fn scan_security_content<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
        _content_types: &ContentTypes,
    ) -> Result<(), ParseError> {
        self.scan_vba_projects(zip, store)?;
        self.scan_ole_objects(zip, store)?;
        self.scan_activex_controls(zip, store)?;
        self.scan_word_external_relationships(zip, store)?;

        Ok(())
    }

    fn scan_vba_projects<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        let vba_paths = [
            "word/vbaProject.bin",
            "xl/vbaProject.bin",
            "ppt/vbaProject.bin",
        ];
        for vba_path in &vba_paths {
            if zip.contains(vba_path) {
                let (mut macro_project, modules) = self.detect_macro_project(zip, vba_path)?;
                for module in modules {
                    let id = module.id;
                    store.insert(IRNode::MacroModule(module));
                    macro_project.modules.push(id);
                }
                store.insert(IRNode::MacroProject(macro_project));
            }
        }
        Ok(())
    }

    fn scan_ole_objects<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        let ole_files: Vec<String> = zip
            .list_prefix("word/embeddings/")
            .into_iter()
            .chain(zip.list_prefix("xl/embeddings/"))
            .chain(zip.list_prefix("ppt/embeddings/"))
            .filter(|p| p.ends_with(".bin") || p.ends_with(".ole"))
            .map(|s| s.to_string())
            .collect();

        for ole_path in ole_files {
            let ole_object = self.detect_ole_object(zip, &ole_path)?;
            store.insert(IRNode::OleObject(ole_object));
        }
        Ok(())
    }

    fn link_shapes_to_shared_parts(&self, store: &mut IrStore) {
        use docir_core::ir::IRNode;
        use docir_core::types::NodeType;

        let mut chart_by_path: HashMap<String, NodeId> = HashMap::new();
        let mut smartart_by_path: HashMap<String, NodeId> = HashMap::new();
        let mut media_by_path: HashMap<String, NodeId> = HashMap::new();
        let mut ole_by_path: HashMap<String, NodeId> = HashMap::new();

        for (id, node) in store.iter() {
            match node {
                IRNode::ChartData(chart) => {
                    if let Some(span) = chart.span.as_ref() {
                        chart_by_path.insert(span.file_path.clone(), *id);
                    }
                }
                IRNode::SmartArtPart(part) => {
                    smartart_by_path.insert(part.path.clone(), *id);
                }
                IRNode::MediaAsset(media) => {
                    media_by_path.insert(media.path.clone(), *id);
                }
                IRNode::OleObject(ole) => {
                    if let Some(span) = ole.span.as_ref() {
                        ole_by_path.insert(span.file_path.clone(), *id);
                    }
                }
                _ => {}
            }
        }

        let shape_ids: Vec<NodeId> = store.iter_ids_by_type(NodeType::Shape).collect();
        for shape_id in shape_ids {
            let Some(IRNode::Shape(shape)) = store.get_mut(shape_id) else {
                continue;
            };
            if let Some(target) = shape.media_target.as_ref() {
                if let Some(id) = resolve_media_asset(&media_by_path, target) {
                    shape.media_asset = Some(id);
                }
                if let Some(id) = chart_by_path.get(target) {
                    shape.chart_id = Some(*id);
                }
                if let Some(id) = ole_by_path.get(target) {
                    shape.ole_object = Some(*id);
                    shape.shape_type = docir_core::ir::ShapeType::OleObject;
                }
            }
            if !shape.related_targets.is_empty() {
                for target in &shape.related_targets {
                    if let Some(id) = smartart_by_path.get(target) {
                        shape.smartart_parts.push(*id);
                    }
                    if let Some(id) = ole_by_path.get(target) {
                        shape.ole_object = Some(*id);
                        shape.shape_type = docir_core::ir::ShapeType::OleObject;
                    }
                }
                shape.smartart_parts.sort_by_key(|id| id.as_u64());
                shape.smartart_parts.dedup();
            }
        }

        let slide_ids: Vec<NodeId> = store.iter_ids_by_type(NodeType::Slide).collect();
        for slide_id in slide_ids {
            let Some(IRNode::Slide(slide)) = store.get_mut(slide_id) else {
                continue;
            };
            for anim in &mut slide.animations {
                if let Some(target) = anim.target.as_ref() {
                    if let Some(id) = resolve_media_asset(&media_by_path, target) {
                        anim.media_asset = Some(id);
                    }
                }
            }
        }
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

        let mut probe = [0u8; 16];
        let read = reader.read(&mut probe)?;
        reader.seek(SeekFrom::Start(0))?;
        let head = &probe[..read];

        if is_rtf_bytes(head) {
            let parser = RtfParser::with_config(self.config.clone());
            return parse_with(parser, reader);
        }

        if is_ole_container(head) {
            let parser = HwpParser::with_config(self.config.clone());
            return parse_with(parser, reader);
        }

        if !is_zip_container(head) {
            return Err(ParseError::UnsupportedFormat(
                "Unknown package format (not OLE/CFB or ZIP)".to_string(),
            ));
        }

        let mut inspector = SecureZipReader::new(&mut reader, self.config.zip_config.clone())?;
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
            let parser = OoxmlParser::with_config(self.config.clone());
            return parse_with(parser, reader);
        }
        if is_hwpx {
            let parser = HwpxParser::with_config(self.config.clone());
            return parse_with(parser, reader);
        }
        if is_odf {
            let parser = OdfParser::with_config(self.config.clone());
            return parse_with(parser, reader);
        }

        Err(ParseError::UnsupportedFormat(
            "Unknown package format (missing [Content_Types].xml and mimetype)".to_string(),
        ))
    }
}

fn is_zip_container(data: &[u8]) -> bool {
    data.len() >= 4 && data[0] == b'P' && data[1] == b'K'
}

fn parse_with<P, R>(parser: P, reader: R) -> Result<ParsedDocument, ParseError>
where
    P: FormatParser,
    R: Read + Seek,
{
    parser.parse_reader(reader)
}

fn resolve_media_asset(media_by_path: &HashMap<String, NodeId>, target: &str) -> Option<NodeId> {
    if let Some(id) = media_by_path.get(target) {
        return Some(*id);
    }
    let trimmed = target.trim_start_matches('/');
    for (path, id) in media_by_path {
        if path.ends_with(trimmed) || trimmed.ends_with(path) {
            return Some(*id);
        }
    }
    None
}

fn find_stream_case<'a>(streams: &'a [String], name: &str) -> Option<&'a str> {
    let target = name.to_ascii_uppercase();
    streams
        .iter()
        .find(|s| s.to_ascii_uppercase() == target)
        .map(|s| s.as_str())
}

fn parse_vba_project_text(
    text: &str,
) -> (
    Option<String>,
    Vec<(String, docir_core::security::MacroModuleType)>,
    Vec<docir_core::security::MacroReference>,
    bool,
) {
    let mut project_name = None;
    let mut modules = Vec::new();
    let mut references = Vec::new();
    let mut protected = false;

    for line in text.lines() {
        let line = line.trim();
        if line.is_empty() {
            continue;
        }
        if line.starts_with("Name=") {
            project_name = Some(
                line.trim_start_matches("Name=")
                    .trim()
                    .trim_matches('"')
                    .to_string(),
            );
        } else if line.starts_with("Module=") {
            let name = line
                .trim_start_matches("Module=")
                .split('/')
                .next()
                .unwrap_or("")
                .to_string();
            if !name.is_empty() {
                modules.push((name, docir_core::security::MacroModuleType::Standard));
            }
        } else if line.starts_with("Class=") {
            let name = line
                .trim_start_matches("Class=")
                .split('/')
                .next()
                .unwrap_or("")
                .to_string();
            if !name.is_empty() {
                modules.push((name, docir_core::security::MacroModuleType::Class));
            }
        } else if line.starts_with("Document=") {
            let name = line
                .trim_start_matches("Document=")
                .split('/')
                .next()
                .unwrap_or("")
                .to_string();
            if !name.is_empty() {
                modules.push((name, docir_core::security::MacroModuleType::Document));
            }
        } else if line.starts_with("Reference=") {
            references.push(docir_core::security::MacroReference {
                name: line.to_string(),
                guid: None,
                path: None,
                major_version: None,
                minor_version: None,
            });
        } else if line.starts_with("DPB=") {
            protected = true;
        }
    }

    (project_name, modules, references, protected)
}

fn vba_decompress(data: &[u8]) -> Option<Vec<u8>> {
    if data.is_empty() {
        return None;
    }
    if data[0] != 0x01 {
        return Some(data.to_vec());
    }

    let mut out = Vec::new();
    let mut pos = 1usize;
    while pos + 2 <= data.len() {
        let header = u16::from_le_bytes([data[pos], data[pos + 1]]);
        pos += 2;
        let chunk_size = ((header & 0x0FFF) as usize) + 3;
        let compressed = (header & 0x8000) != 0;
        if pos + chunk_size > data.len() {
            break;
        }
        if !compressed {
            out.extend_from_slice(&data[pos..pos + chunk_size]);
            pos += chunk_size;
            continue;
        }

        let chunk_end = pos + chunk_size;
        let mut chunk_out = Vec::new();
        while pos < chunk_end {
            let flags = data[pos];
            pos += 1;
            for bit in 0..8 {
                if pos >= chunk_end {
                    break;
                }
                if (flags & (1 << bit)) == 0 {
                    chunk_out.push(data[pos]);
                    pos += 1;
                } else {
                    if pos + 2 > chunk_end {
                        break;
                    }
                    let token = u16::from_le_bytes([data[pos], data[pos + 1]]);
                    pos += 2;
                    let (offset, length) = decode_copy_token(token, chunk_out.len());
                    for _ in 0..length {
                        if offset == 0 || offset > chunk_out.len() {
                            break;
                        }
                        let b = chunk_out[chunk_out.len() - offset];
                        chunk_out.push(b);
                    }
                }
            }
        }
        out.extend_from_slice(&chunk_out);
    }
    Some(out)
}

fn decode_copy_token(token: u16, decompressed_len: usize) -> (usize, usize) {
    let mut bit_count = 0usize;
    let mut val = if decompressed_len == 0 {
        1
    } else {
        decompressed_len
    };
    while val > 0 {
        bit_count += 1;
        val >>= 1;
    }
    let offset_bits = if bit_count < 4 { 4 } else { bit_count };
    let length_bits = 16 - offset_bits;
    let offset_mask = (1u16 << offset_bits) - 1;
    let length_mask = (1u16 << length_bits) - 1;
    let offset = ((token >> length_bits) & offset_mask) as usize + 1;
    let length = (token & length_mask) as usize + 3;
    (offset, length)
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
