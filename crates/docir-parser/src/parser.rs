//! Main OOXML parser orchestrator.

use crate::error::ParseError;
use crate::hwp::{is_hwpx_mimetype, HwpParser, HwpxParser};
use crate::input::{enforce_input_size, parse_from_bytes, parse_from_file, read_all_with_limit};
use crate::odf::OdfParser;
use crate::ole::{is_ole_container, Cfb};
use crate::ooxml::content_types::ContentTypes;
use crate::ooxml::docx::DocxParser;
use crate::ooxml::part_registry;
use crate::ooxml::pptx::PptxParser;
use crate::ooxml::relationships::{rel_type, Relationships, TargetMode};
use crate::ooxml::xlsx::XlsxParser;
use crate::rtf::RtfParser;
use crate::security_utils::parse_dde_instruction;
use crate::zip_handler::{SecureZipReader, ZipConfig};
use docir_core::ir::column_to_letter;
use docir_core::ir::{
    Cell, CellError, CellValue, CustomProperty, DiagnosticEntry, DiagnosticSeverity, Diagnostics,
    Document, DocumentMetadata, ExtensionPart, ExtensionPartKind, IRNode, IrNode as IrNodeTrait,
    MediaAsset, PropertyValue, SheetKind, SheetState, Worksheet,
};
use docir_core::normalize::normalize_store;
use docir_core::security::{
    ExternalRefType, ExternalReference, MacroProject, OleObject, SecurityInfo,
};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read, Seek, SeekFrom};
use std::path::Path;

mod parser_docx;
mod parser_pptx;
mod parser_xlsx;
mod shared_parts;

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
    /// ODF fast-mode threshold (content.xml size in bytes).
    pub odf_fast_threshold_bytes: u64,
    /// Force ODF fast mode regardless of size.
    pub odf_force_fast: bool,
    /// ODF fast-mode sample rows (0 = no sampling).
    pub odf_fast_sample_rows: u32,
    /// ODF fast-mode sample columns (0 = no sampling).
    pub odf_fast_sample_cols: u32,
    /// ODF maximum cells (None = unlimited).
    pub odf_max_cells: Option<u64>,
    /// ODF maximum rows (None = unlimited).
    pub odf_max_rows: Option<u64>,
    /// ODF maximum paragraphs (None = unlimited).
    pub odf_max_paragraphs: Option<u64>,
    /// ODF maximum content.xml bytes (None = unlimited).
    pub odf_max_bytes: Option<u64>,
    /// Enable parallel parsing of ODF sheets when possible.
    pub odf_parallel_sheets: bool,
    /// Max threads for parallel ODF sheet parsing (None = auto).
    pub odf_parallel_max_threads: Option<usize>,
    /// Password for decrypting encrypted ODF parts.
    pub odf_password: Option<String>,
    /// RTF max group depth (0 = unlimited).
    pub rtf_max_group_depth: usize,
    /// RTF max object hex length (0 = unlimited).
    pub rtf_max_object_hex_len: usize,
    /// Force parse encrypted HWP streams.
    pub hwp_force_parse_encrypted: bool,
    /// Password for decrypting encrypted HWP streams.
    pub hwp_password: Option<String>,
    /// Dump HWP stream metadata (hash, size, compression).
    pub hwp_dump_streams: bool,
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
            odf_fast_threshold_bytes: 50 * 1024 * 1024,
            odf_force_fast: false,
            odf_fast_sample_rows: 0,
            odf_fast_sample_cols: 0,
            odf_max_cells: None,
            odf_max_rows: None,
            odf_max_paragraphs: None,
            odf_max_bytes: None,
            odf_parallel_sheets: false,
            odf_parallel_max_threads: None,
            odf_password: None,
            rtf_max_group_depth: 256,
            rtf_max_object_hex_len: 64 * 1024 * 1024,
            hwp_force_parse_encrypted: false,
            hwp_password: None,
            hwp_dump_streams: false,
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

    /// Parses a file from the filesystem.
    pub fn parse_file<P: AsRef<Path>>(&self, path: P) -> Result<ParsedDocument, ParseError> {
        parse_from_file(path, |reader| self.parse_reader(reader))
    }

    /// Parses from a byte slice.
    pub fn parse_bytes(&self, data: &[u8]) -> Result<ParsedDocument, ParseError> {
        parse_from_bytes(data, |reader| self.parse_reader(reader))
    }

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
        let styles_id = self.parse_docx_part_by_rel(
            zip,
            main_part_path,
            doc_rels,
            rel_type::STYLES,
            |part_path, xml| {
                let id = parser.parse_styles(xml).ok()?;
                if let Some(IRNode::StyleSet(set)) = parser.store_mut().get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
            },
        );

        let styles_with_effects_id =
            self.parse_docx_part_by_path(zip, "word/stylesWithEffects.xml", |part_path, xml| {
                let id = parser.parse_styles_with_effects(xml).ok()?;
                if let Some(IRNode::StyleSet(set)) = parser.store_mut().get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
            });

        let numbering_id = self.parse_docx_part_by_rel(
            zip,
            main_part_path,
            doc_rels,
            rel_type::NUMBERING,
            |part_path, xml| {
                let id = parser.parse_numbering(xml).ok()?;
                if let Some(IRNode::NumberingSet(set)) = parser.store_mut().get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
            },
        );

        let comments = doc_rels
            .get_first_by_type(rel_type::COMMENTS)
            .and_then(|rel| {
                let part_path = Relationships::resolve_target(main_part_path, &rel.target);
                let rels = self.read_part_relationships(zip, &part_path);
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
            let rels = self.read_part_relationships(zip, "word/comments.xml");
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
                let rels = self.read_part_relationships(zip, &part_path);
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
                let rels = self.read_part_relationships(zip, &part_path);
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

        let settings_id = self.parse_docx_part_by_rel(
            zip,
            main_part_path,
            doc_rels,
            rel_type::SETTINGS,
            |part_path, xml| {
                let id = parser.parse_settings(xml).ok()?;
                if let Some(IRNode::WordSettings(settings)) = parser.store_mut().get_mut(id) {
                    settings.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
            },
        );

        let web_settings_id = self.parse_docx_part_by_rel(
            zip,
            main_part_path,
            doc_rels,
            rel_type::WEB_SETTINGS,
            |part_path, xml| {
                let id = parser.parse_web_settings(xml).ok()?;
                if let Some(IRNode::WebSettings(settings)) = parser.store_mut().get_mut(id) {
                    settings.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
            },
        );

        let mut font_table_id = self.parse_docx_part_by_rel(
            zip,
            main_part_path,
            doc_rels,
            rel_type::FONT_TABLE,
            |part_path, xml| {
                let id = parser.parse_font_table(xml).ok()?;
                if let Some(IRNode::FontTable(table)) = parser.store_mut().get_mut(id) {
                    table.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
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

        let comments_ext_id =
            self.parse_docx_part_by_path(zip, "word/commentsExtended.xml", |part_path, xml| {
                let id = parser.parse_comments_extended(xml).ok()?;
                if let Some(IRNode::CommentExtensionSet(set)) = parser.store_mut().get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
            });

        let comments_id_map_id =
            self.parse_docx_part_by_path(zip, "word/commentsIds.xml", |part_path, xml| {
                let id = parser.parse_comments_ids(xml).ok()?;
                if let Some(IRNode::CommentIdMap(map)) = parser.store_mut().get_mut(id) {
                    map.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
            });

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
            let rels = self.read_part_relationships(zip, &part_path);
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
            let rels = self.read_part_relationships(zip, &part_path);
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

    fn read_part_relationships<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        part_path: &str,
    ) -> Relationships {
        let rels_path = Self::get_rels_path(part_path);
        if zip.contains(&rels_path) {
            if let Ok(rels_xml) = zip.read_file_string(&rels_path) {
                if let Ok(rels) = Relationships::parse(&rels_xml) {
                    return rels;
                }
            }
        }
        Relationships::default()
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

    fn add_extension_parts_and_diagnostics<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        content_types: &ContentTypes,
        store: &mut IrStore,
        root_id: NodeId,
    ) -> Result<(), ParseError> {
        let mut seen_paths = collect_seen_paths(store);
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
        let coverage_entries = build_coverage_diagnostics(
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

    /// Parse document metadata.
    fn parse_metadata<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
    ) -> Result<Option<NodeId>, ParseError> {
        if zip.contains("docProps/core.xml") {
            let metadata = self.build_metadata(zip);
            Ok(metadata.map(|m| m.id))
        } else {
            Ok(None)
        }
    }

    /// Build metadata from core.xml and app.xml.
    fn build_metadata<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
    ) -> Option<DocumentMetadata> {
        let mut metadata = DocumentMetadata::new();

        // Parse core.xml (Dublin Core properties)
        if let Ok(core_xml) = zip.read_file_string("docProps/core.xml") {
            self.parse_core_properties(&core_xml, &mut metadata);
        }

        // Parse app.xml (application properties)
        if let Ok(app_xml) = zip.read_file_string("docProps/app.xml") {
            self.parse_app_properties(&app_xml, &mut metadata);
        }

        // Parse custom.xml (custom properties)
        if let Ok(custom_xml) = zip.read_file_string("docProps/custom.xml") {
            self.parse_custom_properties(&custom_xml, &mut metadata);
        }

        Some(metadata)
    }

    /// Parse core.xml properties.
    fn parse_core_properties(&self, xml: &str, metadata: &mut DocumentMetadata) {
        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_element = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    current_element = String::from_utf8_lossy(e.name().as_ref()).to_string();
                }
                Ok(Event::Text(e)) => {
                    let text = e.unescape().unwrap_or_default().to_string();
                    match current_element.as_str() {
                        "dc:title" | "title" => metadata.title = Some(text),
                        "dc:subject" | "subject" => metadata.subject = Some(text),
                        "dc:creator" | "creator" => metadata.creator = Some(text),
                        "cp:keywords" | "keywords" => metadata.keywords = Some(text),
                        "dc:description" | "description" => metadata.description = Some(text),
                        "cp:lastModifiedBy" | "lastModifiedBy" => {
                            metadata.last_modified_by = Some(text)
                        }
                        "cp:revision" | "revision" => metadata.revision = Some(text),
                        "dcterms:created" | "created" => metadata.created = Some(text),
                        "dcterms:modified" | "modified" => metadata.modified = Some(text),
                        "cp:category" | "category" => metadata.category = Some(text),
                        "cp:contentStatus" | "contentStatus" => {
                            metadata.content_status = Some(text)
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(_)) => {
                    current_element.clear();
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }
    }

    /// Parse app.xml properties.
    fn parse_app_properties(&self, xml: &str, metadata: &mut DocumentMetadata) {
        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_element = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    current_element = String::from_utf8_lossy(e.name().as_ref()).to_string();
                }
                Ok(Event::Text(e)) => {
                    let text = e.unescape().unwrap_or_default().to_string();
                    match current_element.as_str() {
                        "Application" => metadata.application = Some(text),
                        "AppVersion" => metadata.app_version = Some(text),
                        "Company" => metadata.company = Some(text),
                        "Manager" => metadata.manager = Some(text),
                        _ => {}
                    }
                }
                Ok(Event::End(_)) => {
                    current_element.clear();
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }
    }

    /// Parse custom properties from custom.xml.
    fn parse_custom_properties(&self, xml: &str, metadata: &mut DocumentMetadata) {
        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_prop: Option<CustomProperty> = None;
        let mut current_value_tag: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                    if name.ends_with("property") {
                        let mut prop = CustomProperty {
                            name: String::new(),
                            value: PropertyValue::String(String::new()),
                            format_id: None,
                            property_id: None,
                        };
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"name" => {
                                    prop.name = String::from_utf8_lossy(&attr.value).to_string()
                                }
                                b"fmtid" => {
                                    prop.format_id =
                                        Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"pid" => {
                                    prop.property_id =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                                }
                                _ => {}
                            }
                        }
                        current_prop = Some(prop);
                    } else if name.starts_with("vt:") || name.contains(":") {
                        current_value_tag = Some(name);
                    }
                }
                Ok(Event::Text(e)) => {
                    if let Some(tag) = &current_value_tag {
                        if let Some(prop) = current_prop.as_mut() {
                            let text = e.unescape().unwrap_or_default().to_string();
                            prop.value = match tag.as_str() {
                                "vt:lpwstr" | "vt:lpstr" | "vt:bstr" => PropertyValue::String(text),
                                "vt:i2" | "vt:i4" | "vt:int" | "vt:integer" => {
                                    PropertyValue::Integer(text.parse::<i64>().unwrap_or(0))
                                }
                                "vt:r4" | "vt:r8" | "vt:float" => {
                                    PropertyValue::Float(text.parse::<f64>().unwrap_or(0.0))
                                }
                                "vt:bool" => PropertyValue::Boolean(text == "true" || text == "1"),
                                "vt:filetime" => PropertyValue::DateTime(text),
                                "vt:blob" => PropertyValue::Blob(text),
                                _ => PropertyValue::String(text),
                            };
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                    if name.ends_with("property") {
                        if let Some(prop) = current_prop.take() {
                            metadata.custom_properties.push(prop);
                        }
                    } else if current_value_tag.as_deref() == Some(&name) {
                        current_value_tag = None;
                    }
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }
    }

    /// Scan for security-relevant content.
    fn scan_security_content<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
        root_id: NodeId,
        _content_types: &ContentTypes,
    ) -> Result<(), ParseError> {
        // Check for VBA project
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
                let macro_id = macro_project.id;
                store.insert(IRNode::MacroProject(macro_project));

                if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
                    doc.security.macro_project = Some(macro_id);
                }
            }
        }

        // Check for OLE objects
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
            let ole_id = ole_object.id;
            store.insert(IRNode::OleObject(ole_object));

            if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
                doc.security.ole_objects.push(ole_id);
            }
        }

        // ActiveX controls
        let mut activex_bin_seen: HashSet<String> = HashSet::new();
        let activex_paths: Vec<String> = zip
            .list_prefix("word/activeX/")
            .into_iter()
            .chain(zip.list_prefix("xl/activeX/"))
            .chain(zip.list_prefix("ppt/activeX/"))
            .filter(|p| p.ends_with(".xml"))
            .map(|s| s.to_string())
            .collect();
        for path in activex_paths {
            let xml = zip.read_file_string(&path)?;
            if let Some(mut control) = parse_activex_xml(&xml, &path) {
                control.span = Some(SourceSpan::new(&path));
                let id = control.id;
                store.insert(IRNode::ActiveXControl(control));
                if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
                    doc.security.activex_controls.push(id);
                }
            }

            let rels_path = Self::get_rels_path(&path);
            if zip.contains(&rels_path) {
                if let Ok(rels_xml) = zip.read_file_string(&rels_path) {
                    if let Ok(rels) = Relationships::parse(&rels_xml) {
                        for rel in rels.by_id.values() {
                            if !rel.target.ends_with(".bin")
                                && !rel.rel_type.contains("activeXControlBinary")
                            {
                                continue;
                            }
                            let bin_path = Relationships::resolve_target(&path, &rel.target);
                            if activex_bin_seen.insert(bin_path.clone()) && zip.contains(&bin_path)
                            {
                                let ole_object = self.detect_ole_object(zip, &bin_path)?;
                                let ole_id = ole_object.id;
                                store.insert(IRNode::OleObject(ole_object));
                                if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
                                    doc.security.ole_objects.push(ole_id);
                                }
                            }
                        }
                    }
                }
            }
        }

        // Word field DDE instructions
        let mut dde_fields = Vec::new();
        for node in store.values() {
            if let IRNode::Field(field) = node {
                if let Some(instr) = &field.instruction {
                    if let Some(dde) = parse_dde_instruction(instr) {
                        dde_fields.push(dde);
                    }
                }
            }
        }
        if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
            doc.security.dde_fields.extend(dde_fields);
        }

        // Word external relationships
        let rel_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.starts_with("word/") && p.ends_with(".rels"))
            .map(|s| s.to_string())
            .collect();
        for rel_path in rel_paths {
            let rels_xml = zip.read_file_string(&rel_path)?;
            let rels = Relationships::parse(&rels_xml)?;
            for rel in rels.by_id.values() {
                if rel.target_mode == TargetMode::External {
                    let ref_type = match rel.rel_type.as_str() {
                        rel_type::HYPERLINK => ExternalRefType::Hyperlink,
                        rel_type::IMAGE => ExternalRefType::Image,
                        rel_type::OLE_OBJECT => ExternalRefType::OleLink,
                        rel_type::ATTACHED_TEMPLATE => ExternalRefType::AttachedTemplate,
                        _ => ExternalRefType::Other,
                    };
                    let mut ext_ref = ExternalReference::new(ref_type, &rel.target);
                    ext_ref.relationship_id = Some(rel.id.clone());
                    ext_ref.relationship_type = Some(rel.rel_type.clone());
                    ext_ref.span = Some(SourceSpan::new(&rel_path));
                    let id = ext_ref.id;
                    store.insert(IRNode::ExternalReference(ext_ref));
                    if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
                        doc.security.external_refs.push(id);
                    }
                }
            }
        }

        Ok(())
    }

    /// Detect and extract macro project information.
    fn detect_macro_project<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        path: &str,
    ) -> Result<(MacroProject, Vec<docir_core::security::MacroModule>), ParseError> {
        let data = zip.read_file(path)?;

        let mut project = MacroProject::new();
        project.span = Some(SourceSpan::new(path));

        let cfb = Cfb::parse(data)?;
        let streams = cfb.list_streams();

        let project_stream =
            find_stream_case(&streams, "VBA/PROJECT").and_then(|p| cfb.read_stream(p));
        let project_text = project_stream
            .as_ref()
            .map(|data| String::from_utf8_lossy(data).to_string())
            .unwrap_or_default();

        let (project_name, module_defs, references, is_protected) =
            parse_vba_project_text(&project_text);
        project.name = project_name;
        project.references = references;
        project.is_protected = is_protected;

        let mut auto_exec = Vec::new();
        let mut modules_out = Vec::new();
        for (module_name, module_type) in module_defs {
            let stream_path = format!("VBA/{module_name}");
            let data = find_stream_case(&streams, &stream_path)
                .and_then(|p| cfb.read_stream(p))
                .and_then(|raw| vba_decompress(&raw));

            let mut module =
                docir_core::security::MacroModule::new(module_name.clone(), module_type);
            module.span = Some(SourceSpan::new(path));

            if let Some(src) = data {
                let source = String::from_utf8_lossy(&src).to_string();
                let (procedures, suspicious) = analyze_vba_source(&source);
                for proc_name in &procedures {
                    if docir_core::security::AUTO_EXEC_PROCEDURES
                        .iter()
                        .any(|p| p.eq_ignore_ascii_case(proc_name))
                    {
                        auto_exec.push(proc_name.clone());
                    }
                }
                module.procedures = procedures;
                module.suspicious_calls = suspicious;

                if self.config.extract_macro_source {
                    module.source_code = Some(source);
                }
                if self.config.compute_hashes {
                    use sha2::Digest;
                    let mut hasher = sha2::Sha256::new();
                    hasher.update(&src);
                    module.source_hash = Some(hex::encode(hasher.finalize()));
                }
            }

            modules_out.push(module);
        }

        if !auto_exec.is_empty() {
            project.has_auto_exec = true;
            project.auto_exec_procedures = auto_exec;
        }

        Ok((project, modules_out))
    }

    /// Detect and extract basic OLE object information.
    fn detect_ole_object<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        path: &str,
    ) -> Result<OleObject, ParseError> {
        let data = zip.read_file(path)?;

        let mut ole = OleObject::new();
        ole.span = Some(SourceSpan::new(path));
        ole.size_bytes = data.len() as u64;

        // Compute hash if configured
        if self.config.compute_hashes {
            use sha2::{Digest, Sha256};
            let mut hasher = Sha256::new();
            hasher.update(&data);
            let hash = hasher.finalize();
            ole.data_hash = Some(hex::encode(hash));
        }

        Ok(ole)
    }

    /// Get the relationships file path for a given part.
    fn get_rels_path(part_path: &str) -> String {
        if let Some(idx) = part_path.rfind('/') {
            let dir = &part_path[..idx + 1];
            let file = &part_path[idx + 1..];
            format!("{}_rels/{}.rels", dir, file)
        } else {
            format!("_rels/{}.rels", part_path)
        }
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

    /// Parses a file from the filesystem.
    pub fn parse_file<P: AsRef<Path>>(&self, path: P) -> Result<ParsedDocument, ParseError> {
        parse_from_file(path, |reader| self.parse_reader(reader))
    }

    /// Parses from a byte slice.
    pub fn parse_bytes(&self, data: &[u8]) -> Result<ParsedDocument, ParseError> {
        parse_from_bytes(data, |reader| self.parse_reader(reader))
    }

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
            return parser.parse_reader(reader);
        }

        if is_ole_container(head) {
            let parser = HwpParser::with_config(self.config.clone());
            return parser.parse_reader(reader);
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
            return parser.parse_reader(reader);
        }
        if is_hwpx {
            let parser = HwpxParser::with_config(self.config.clone());
            return parser.parse_reader(reader);
        }
        if is_odf {
            let parser = OdfParser::with_config(self.config.clone());
            return parser.parse_reader(reader);
        }

        Err(ParseError::UnsupportedFormat(
            "Unknown package format (missing [Content_Types].xml and mimetype)".to_string(),
        ))
    }
}

fn is_zip_container(data: &[u8]) -> bool {
    data.len() >= 4 && data[0] == b'P' && data[1] == b'K'
}

fn is_rtf_bytes(data: &[u8]) -> bool {
    data.starts_with(b"{\\rtf")
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

fn collect_seen_paths(store: &IrStore) -> HashSet<String> {
    let mut seen = HashSet::new();
    for node in store.values() {
        if let Some(span) = node.source_span() {
            seen.insert(span.file_path.clone());
        }
        if let IRNode::MediaAsset(media) = node {
            seen.insert(media.path.clone());
        }
        if let IRNode::CustomXmlPart(part) = node {
            seen.insert(part.path.clone());
        }
        if let IRNode::ExtensionPart(part) = node {
            seen.insert(part.path.clone());
        }
        if let IRNode::RelationshipGraph(graph) = node {
            seen.insert(graph.source.clone());
        }
    }
    seen
}

fn build_coverage_diagnostics(
    format: DocumentFormat,
    content_types: &ContentTypes,
    all_paths: Vec<String>,
    seen_paths: &HashSet<String>,
) -> Vec<DiagnosticEntry> {
    let all_paths: Vec<String> = all_paths
        .into_iter()
        .filter(|p| !p.ends_with('/') && !p.starts_with("[trash]/"))
        .collect();
    let registry = part_registry::registry_for(format);
    let mut entries = Vec::new();

    let mut matched_parts = 0usize;
    let mut complete = 0usize;
    let mut pending = 0usize;
    let mut missing = 0usize;
    let mut by_status: HashMap<&'static str, usize> = HashMap::new();

    for spec in registry {
        let mut matched = false;
        for path in &all_paths {
            let ct = content_types.get_content_type(path);
            if spec.matches(path, ct) {
                matched = true;
                matched_parts += 1;
                let status = if seen_paths.contains(path) {
                    complete += 1;
                    "complete"
                } else {
                    pending += 1;
                    "pending"
                };
                *by_status.entry(status).or_insert(0) += 1;
                let ct_str = ct.unwrap_or("unknown");
                let message = format!(
                    "{} part: {} (content-type={}, parser={})",
                    status, path, ct_str, spec.expected_parser
                );
                entries.push(DiagnosticEntry {
                    severity: if status == "complete" {
                        DiagnosticSeverity::Info
                    } else {
                        DiagnosticSeverity::Warning
                    },
                    code: "COVERAGE_PART".to_string(),
                    message: message.clone(),
                    path: Some(path.clone()),
                });
                entries.push(DiagnosticEntry {
                    severity: DiagnosticSeverity::Info,
                    code: "COVERAGE_PART_ROW".to_string(),
                    message,
                    path: Some(path.clone()),
                });
            }
        }
        if !matched {
            missing += 1;
            entries.push(DiagnosticEntry {
                severity: DiagnosticSeverity::Warning,
                code: "COVERAGE_MISSING".to_string(),
                message: format!(
                    "missing part for pattern {} (expected parser={})",
                    spec.pattern, spec.expected_parser
                ),
                path: Some(spec.pattern.to_string()),
            });
        }
    }

    let mut inventory: Vec<(String, String)> = Vec::new();
    let mut unknown_ct_paths: Vec<String> = Vec::new();
    let mut ct_histogram: HashMap<String, usize> = HashMap::new();
    for path in &all_paths {
        if let Some(ct) = content_types.get_content_type(path) {
            inventory.push((path.clone(), ct.to_string()));
            *ct_histogram.entry(ct.to_string()).or_insert(0) += 1;
        } else {
            unknown_ct_paths.push(path.clone());
        }
    }
    inventory.sort_by(|a, b| a.0.cmp(&b.0));
    for (path, ct) in inventory {
        entries.push(DiagnosticEntry {
            severity: DiagnosticSeverity::Info,
            code: "CONTENT_TYPE_INVENTORY".to_string(),
            message: format!("content-type: {} => {}", path, ct),
            path: Some(path),
        });
    }
    unknown_ct_paths.sort();
    for path in &unknown_ct_paths {
        entries.push(DiagnosticEntry {
            severity: DiagnosticSeverity::Warning,
            code: "CONTENT_TYPE_UNKNOWN".to_string(),
            message: format!("no content-type entry for path {}", path),
            path: Some(path.clone()),
        });
    }

    entries.push(DiagnosticEntry {
        severity: DiagnosticSeverity::Info,
        code: "COVERAGE_SUMMARY".to_string(),
        message: format!(
            "coverage summary: matched={}, complete={}, pending={}, missing_patterns={}, unknown_content_types={}",
            matched_parts, complete, pending, missing, unknown_ct_paths.len()
        ),
        path: None,
    });
    entries.push(DiagnosticEntry {
        severity: DiagnosticSeverity::Info,
        code: "COVERAGE_COUNTS".to_string(),
        message: format!(
            "counts: complete={}, pending={}, missing_patterns={}, unknown_content_types={}",
            complete,
            pending,
            missing,
            unknown_ct_paths.len()
        ),
        path: None,
    });

    let mut hist: Vec<(String, usize)> = ct_histogram.into_iter().collect();
    hist.sort_by(|a, b| a.0.cmp(&b.0));
    for (ct, count) in hist {
        entries.push(DiagnosticEntry {
            severity: DiagnosticSeverity::Info,
            code: "CONTENT_TYPE_HISTOGRAM".to_string(),
            message: format!("content-type-count: {} => {}", ct, count),
            path: Some(ct),
        });
    }

    entries
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

fn analyze_vba_source(source: &str) -> (Vec<String>, Vec<docir_core::security::SuspiciousCall>) {
    let mut procedures = Vec::new();
    let mut suspicious = Vec::new();

    for (idx, line) in source.lines().enumerate() {
        let raw = line.trim();
        if raw.is_empty() {
            continue;
        }

        let lower = raw.to_ascii_lowercase();
        let mut tokens: Vec<&str> = raw.split_whitespace().collect();
        if tokens.len() >= 2 {
            if tokens[0].eq_ignore_ascii_case("private")
                || tokens[0].eq_ignore_ascii_case("public")
                || tokens[0].eq_ignore_ascii_case("friend")
                || tokens[0].eq_ignore_ascii_case("static")
            {
                tokens.remove(0);
            }
        }
        if tokens.len() >= 2 {
            let keyword = tokens[0].to_ascii_lowercase();
            if keyword == "sub" || keyword == "function" {
                let name = tokens[1].split('(').next().unwrap_or("").to_string();
                if !name.is_empty() {
                    procedures.push(name);
                }
            }
        }

        for &(call, category) in docir_core::security::SUSPICIOUS_VBA_CALLS {
            if lower.contains(&call.to_ascii_lowercase()) {
                suspicious.push(docir_core::security::SuspiciousCall {
                    name: call.to_string(),
                    category,
                    line: Some((idx + 1) as u32),
                });
            }
        }
    }

    (procedures, suspicious)
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

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(pos) => &name[pos + 1..],
        None => name,
    }
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
mod tests {
    use super::*;
    use docir_core::types::NodeType;
    use docir_core::IrSummary;
    use std::fs::File;
    use std::io::Write;
    use std::time::{SystemTime, UNIX_EPOCH};
    use zip::write::FileOptions;

    #[test]
    fn test_parse_custom_properties() {
        let xml = r#"
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
                    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
          <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="AuthorId">
            <vt:i4>123</vt:i4>
          </property>
          <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Flag">
            <vt:bool>true</vt:bool>
          </property>
          <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Title">
            <vt:lpwstr>Doc</vt:lpwstr>
          </property>
        </Properties>
        "#;
        let mut metadata = DocumentMetadata::new();
        let parser = OoxmlParser::new();
        parser.parse_custom_properties(xml, &mut metadata);
        assert_eq!(metadata.custom_properties.len(), 3);
        assert!(matches!(
            metadata.custom_properties[0].value,
            PropertyValue::Integer(123)
        ));
        assert!(matches!(
            metadata.custom_properties[1].value,
            PropertyValue::Boolean(true)
        ));
        assert!(matches!(
            metadata.custom_properties[2].value,
            PropertyValue::String(_)
        ));
    }

    #[test]
    fn test_parse_minimal_docx() {
        let path = create_minimal_docx(true);
        let parser = OoxmlParser::new();
        let parsed = parser.parse_file(&path).expect("parse minimal docx");
        assert!(parsed.document().is_some());
        std::fs::remove_file(path).ok();
    }

    #[test]
    fn test_missing_main_part_fails_validation() {
        let path = create_minimal_docx(false);
        let parser = OoxmlParser::new();
        let err = parser
            .parse_file(&path)
            .expect_err("missing part should fail");
        match err {
            ParseError::MissingPart(_) => {}
            other => panic!("unexpected error: {other:?}"),
        }
        std::fs::remove_file(path).ok();
    }

    #[test]
    fn test_docx_paragraph_and_run_properties() {
        let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:pPr>
                <w:jc w:val="center"/>
                <w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>
                <w:ind w:left="720" w:firstLine="360"/>
                <w:outlineLvl w:val="2"/>
              </w:pPr>
              <w:r>
                <w:rPr>
                  <w:b/>
                  <w:i/>
                  <w:color w:val="FF0000"/>
                  <w:highlight w:val="yellow"/>
                  <w:sz w:val="24"/>
                </w:rPr>
                <w:t>Hello</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>"#;
        let path = create_docx_with_body(body);
        let parser = OoxmlParser::new();
        let parsed = parser.parse_file(&path).expect("parse docx");

        let mut para_props = None;
        let mut run_props = None;
        for node in parsed.store.values() {
            if let IRNode::Paragraph(p) = node {
                para_props = Some(p.properties.clone());
            } else if let IRNode::Run(r) = node {
                run_props = Some(r.properties.clone());
            }
        }

        let para_props = para_props.expect("paragraph props");
        assert!(matches!(
            para_props.alignment,
            Some(docir_core::ir::TextAlignment::Center)
        ));
        assert_eq!(
            para_props.spacing.as_ref().and_then(|s| s.before),
            Some(120)
        );
        assert_eq!(para_props.spacing.as_ref().and_then(|s| s.after), Some(240));
        assert_eq!(para_props.spacing.as_ref().and_then(|s| s.line), Some(360));
        assert_eq!(
            para_props.indentation.as_ref().and_then(|i| i.left),
            Some(720)
        );
        assert_eq!(
            para_props.indentation.as_ref().and_then(|i| i.first_line),
            Some(360)
        );
        assert_eq!(para_props.outline_level, Some(2));

        let run_props = run_props.expect("run props");
        assert_eq!(run_props.bold, Some(true));
        assert_eq!(run_props.italic, Some(true));
        assert_eq!(run_props.color.as_deref(), Some("FF0000"));
        assert_eq!(run_props.highlight.as_deref(), Some("yellow"));
        assert_eq!(run_props.font_size, Some(24));

        std::fs::remove_file(path).ok();
    }

    #[test]
    fn test_docx_sections_with_headers_and_footers() {
        let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p>
              <w:pPr>
                <w:sectPr>
                  <w:headerReference w:type="default" r:id="rIdHeader1"/>
                  <w:footerReference w:type="default" r:id="rIdFooter1"/>
                  <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
                </w:sectPr>
              </w:pPr>
              <w:r><w:t>Section1</w:t></w:r>
            </w:p>
            <w:p><w:r><w:t>Section2</w:t></w:r></w:p>
            <w:sectPr>
              <w:headerReference w:type="default" r:id="rIdHeader2"/>
            </w:sectPr>
          </w:body>
        </w:document>"#;

        let header1 = r#"
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:p><w:r><w:t>Header1</w:t></w:r></w:p>
        </w:hdr>"#;
        let footer1 = r#"
        <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:p><w:r><w:t>Footer1</w:t></w:r></w:p>
        </w:ftr>"#;
        let header2 = r#"
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:p><w:r><w:t>Header2</w:t></w:r></w:p>
        </w:hdr>"#;

        let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdHeader1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
            Target="header1.xml"/>
          <Relationship Id="rIdFooter1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
            Target="footer1.xml"/>
          <Relationship Id="rIdHeader2"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
            Target="header2.xml"/>
        </Relationships>"#;

        let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/header1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
          <Override PartName="/word/footer1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
          <Override PartName="/word/header2.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
        </Types>"#;

        let path = create_docx_with_relationships(
            body,
            rels,
            content_types,
            &[
                ("word/header1.xml", header1),
                ("word/footer1.xml", footer1),
                ("word/header2.xml", header2),
            ],
        );

        let parser = OoxmlParser::new();
        let parsed = parser.parse_file(&path).expect("parse docx");
        let doc = parsed.document().expect("doc");
        assert_eq!(doc.content.len(), 2);

        let mut section_headers = Vec::new();
        for section_id in &doc.content {
            if let Some(IRNode::Section(section)) = parsed.store.get(*section_id) {
                section_headers.push(section.headers.len());
            }
        }
        assert_eq!(section_headers, vec![1, 1]);

        std::fs::remove_file(path).ok();
    }

    #[test]
    fn test_docx_content_controls_inline_and_block() {
        let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:sdt>
              <w:sdtPr>
                <w:alias w:val="BlockControl"/>
                <w:tag w:val="block-tag"/>
              </w:sdtPr>
              <w:sdtContent>
                <w:p><w:r><w:t>Block content</w:t></w:r></w:p>
              </w:sdtContent>
            </w:sdt>
            <w:p>
              <w:r><w:t>Before</w:t></w:r>
              <w:sdt>
                <w:sdtPr>
                  <w:alias w:val="InlineControl"/>
                  <w:tag w:val="inline-tag"/>
                </w:sdtPr>
                <w:sdtContent>
                  <w:r><w:t>Inline content</w:t></w:r>
                </w:sdtContent>
              </w:sdt>
              <w:r><w:t>After</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>"#;

        let path = create_docx_with_body(body);
        let parser = OoxmlParser::new();
        let parsed = parser.parse_file(&path).expect("parse docx");

        let mut controls = Vec::new();
        for node in parsed.store.values() {
            if let IRNode::ContentControl(ctrl) = node {
                controls.push((ctrl.alias.clone(), ctrl.tag.clone(), ctrl.content.len()));
            }
        }
        controls.sort_by(|a, b| a.0.cmp(&b.0));

        assert_eq!(controls.len(), 2);
        assert_eq!(controls[0].0.as_deref(), Some("BlockControl"));
        assert_eq!(controls[0].1.as_deref(), Some("block-tag"));
        assert!(controls[0].2 >= 1);
        assert_eq!(controls[1].0.as_deref(), Some("InlineControl"));
        assert_eq!(controls[1].1.as_deref(), Some("inline-tag"));
        assert!(controls[1].2 >= 1);

        std::fs::remove_file(path).ok();
    }

    #[test]
    fn test_pptx_animation_media_asset_link() {
        let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld><p:spTree/></p:cSld>
          <p:timing>
            <p:audio r:link="rIdAudio" dur="5000"/>
          </p:timing>
        </p:sld>
        "#;
        let slide_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdAudio"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
            Target="../media/audio1.wav"/>
        </Relationships>
        "#;

        let path = create_pptx_with_media(slide_xml, slide_rels);
        let parser = OoxmlParser::new();
        let parsed = parser.parse_file(&path).expect("parse pptx");
        let doc = parsed.document().expect("doc");
        let slide_id = doc.content[0];
        let slide = match parsed.store.get(slide_id) {
            Some(IRNode::Slide(s)) => s,
            _ => panic!("missing slide"),
        };
        assert_eq!(slide.animations.len(), 1);
        let anim = &slide.animations[0];
        let media_id = anim.media_asset.expect("media asset link");
        let media = match parsed.store.get(media_id) {
            Some(IRNode::MediaAsset(m)) => m,
            _ => panic!("missing media asset"),
        };
        assert_eq!(media.path, "ppt/media/audio1.wav");
        std::fs::remove_file(path).ok();
    }

    #[test]
    fn test_docx_chart_shape_linkage() {
        let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p>
              <w:r>
                <w:drawing>
                  <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                      <c:chart r:id="rIdChart1"/>
                    </a:graphicData>
                  </a:graphic>
                </w:drawing>
              </w:r>
            </w:p>
          </w:body>
        </w:document>
        "#;

        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdChart1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            Target="charts/chart1.xml"/>
        </Relationships>
        "#;

        let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/charts/chart1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
        </Types>"#;

        let chart_xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart>
            <c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:barChart/>
          </c:chart>
        </c:chartSpace>
        "#;

        let path = create_docx_with_relationships(
            body,
            rels_xml,
            content_types,
            &[("word/charts/chart1.xml", chart_xml)],
        );

        let parser = OoxmlParser::new();
        let parsed = parser.parse_file(&path).expect("parse docx");
        let mut shape_chart_ids = Vec::new();
        for node in parsed.store.values() {
            if let IRNode::Shape(shape) = node {
                if let Some(chart_id) = shape.chart_id {
                    shape_chart_ids.push(chart_id);
                }
            }
        }
        assert_eq!(shape_chart_ids.len(), 1);
        let chart = match parsed.store.get(shape_chart_ids[0]) {
            Some(IRNode::ChartData(c)) => c,
            _ => panic!("missing chart node"),
        };
        assert_eq!(chart.title.as_deref(), Some("Sales"));

        std::fs::remove_file(path).ok();
    }

    fn create_minimal_docx(include_document: bool) -> std::path::PathBuf {
        let mut path = std::env::temp_dir();
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_nanos();
        path.push(format!("docir_minimal_{nanos}.docx"));

        let file = File::create(&path).expect("create temp docx");
        let mut zip = zip::ZipWriter::new(file);
        let options = FileOptions::<()>::default();

        let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>"#;

        let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;

        let document = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
          </w:body>
        </w:document>"#;

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.trim().as_bytes()).unwrap();
        zip.add_directory("_rels/", options).unwrap();
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(rels.trim().as_bytes()).unwrap();
        zip.add_directory("word/", options).unwrap();
        if include_document {
            zip.start_file("word/document.xml", options).unwrap();
            zip.write_all(document.trim().as_bytes()).unwrap();
        }
        zip.finish().unwrap();
        path
    }

    fn create_docx_with_body(body_xml: &str) -> std::path::PathBuf {
        let mut path = std::env::temp_dir();
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_nanos();
        path.push(format!("docir_body_{nanos}.docx"));

        let file = File::create(&path).expect("create temp docx");
        let mut zip = zip::ZipWriter::new(file);
        let options = FileOptions::<()>::default();

        let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>"#;

        let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.trim().as_bytes()).unwrap();
        zip.add_directory("_rels/", options).unwrap();
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(rels.trim().as_bytes()).unwrap();
        zip.add_directory("word/", options).unwrap();
        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(body_xml.trim().as_bytes()).unwrap();
        zip.finish().unwrap();
        path
    }

    fn create_docx_with_relationships(
        body_xml: &str,
        rels_xml: &str,
        content_types: &str,
        extra_files: &[(&str, &str)],
    ) -> std::path::PathBuf {
        let mut path = std::env::temp_dir();
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_nanos();
        path.push(format!("docir_docx_{nanos}.docx"));

        let file = File::create(&path).expect("create temp docx");
        let mut zip = zip::ZipWriter::new(file);
        let options = FileOptions::<()>::default();

        let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.trim().as_bytes()).unwrap();
        zip.add_directory("_rels/", options).unwrap();
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(rels.trim().as_bytes()).unwrap();
        zip.add_directory("word/", options).unwrap();
        zip.add_directory("word/_rels/", options).unwrap();
        zip.start_file("word/_rels/document.xml.rels", options)
            .unwrap();
        zip.write_all(rels_xml.trim().as_bytes()).unwrap();
        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(body_xml.trim().as_bytes()).unwrap();

        for (path, xml) in extra_files {
            zip.start_file(*path, options).unwrap();
            zip.write_all(xml.trim().as_bytes()).unwrap();
        }

        zip.finish().unwrap();
        path
    }

    fn build_odf_zip_custom(
        mimetype: &str,
        content_xml: &str,
        manifest_xml: &str,
        extra_files: &[(&str, &[u8])],
    ) -> Vec<u8> {
        let mut buffer = Vec::new();
        let cursor = std::io::Cursor::new(&mut buffer);
        let mut zip = zip::ZipWriter::new(cursor);
        let stored =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
        zip.start_file("mimetype", stored).unwrap();
        zip.write_all(mimetype.as_bytes()).unwrap();

        zip.start_file("META-INF/manifest.xml", FileOptions::<()>::default())
            .unwrap();
        zip.write_all(manifest_xml.as_bytes()).unwrap();

        zip.start_file("content.xml", FileOptions::<()>::default())
            .unwrap();
        zip.write_all(content_xml.as_bytes()).unwrap();

        zip.start_file("meta.xml", FileOptions::<()>::default())
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
  <office:meta>
    <dc:title>Parity</dc:title>
  </office:meta>
</office:document-meta>
"#,
        )
        .unwrap();

        for (path, bytes) in extra_files {
            zip.start_file(*path, FileOptions::<()>::default()).unwrap();
            zip.write_all(bytes).unwrap();
        }

        zip.finish().unwrap();
        buffer
    }

    #[test]
    fn test_parity_docx_odt_core_nodes() {
        let docx_body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
            <w:tbl>
              <w:tr>
                <w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>
              </w:tr>
            </w:tbl>
            <w:p>
              <w:hyperlink r:id="rId3"><w:r><w:t>Link</w:t></w:r></w:hyperlink>
            </w:p>
            <w:p>
              <w:commentRangeStart w:id="0"/>
              <w:r><w:t>Commented</w:t></w:r>
              <w:commentRangeEnd w:id="0"/>
            </w:p>
            <w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>
            <w:p>
              <w:ins w:author="A" w:date="2024-01-01T00:00:00Z">
                <w:r><w:t>Inserted</w:t></w:r>
              </w:ins>
            </w:p>
          </w:body>
        </w:document>
        "#;

        let docx_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId3"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            Target="https://example.com" TargetMode="External"/>
          <Relationship Id="rId4"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            Target="comments.xml"/>
          <Relationship Id="rId5"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
            Target="footnotes.xml"/>
        </Relationships>
        "#;

        let docx_content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/comments.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
          <Override PartName="/word/footnotes.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
        </Types>
        "#;

        let comments_xml = r#"
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:comment w:id="0" w:author="Bob" w:date="2024-01-01T00:00:00Z">
            <w:p><w:r><w:t>Comment text</w:t></w:r></w:p>
          </w:comment>
        </w:comments>
        "#;

        let footnotes_xml = r#"
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:footnote w:id="1">
            <w:p><w:r><w:t>Footnote</w:t></w:r></w:p>
          </w:footnote>
        </w:footnotes>
        "#;

        let docx_path = create_docx_with_relationships(
            docx_body,
            docx_rels,
            docx_content_types,
            &[
                ("word/comments.xml", comments_xml),
                ("word/footnotes.xml", footnotes_xml),
            ],
        );

        let odt_content = r#"
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <text:p>Hello</text:p>
      <table:table>
        <table:table-row>
          <table:table-cell><text:p>Cell</text:p></table:table-cell>
        </table:table-row>
      </table:table>
      <text:p><text:a xlink:href="https://example.com">Link</text:a></text:p>
      <office:annotation>
        <text:p>Comment text</text:p>
      </office:annotation>
      <text:note text:note-class="footnote">
        <text:note-body><text:p>Footnote</text:p></text:note-body>
      </text:note>
      <text:tracked-changes>
        <text:changed-region>
          <text:insertion>
            <text:p>Inserted</text:p>
          </text:insertion>
        </text:changed-region>
      </text:tracked-changes>
    </office:text>
  </office:body>
</office:document-content>
        "#;

        let odt_manifest = r#"
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
        "#;

        let odt_zip = build_odf_zip_custom(
            "application/vnd.oasis.opendocument.text",
            odt_content,
            odt_manifest,
            &[],
        );

        let parser = DocumentParser::new();
        let docx = parser.parse_file(&docx_path).expect("docx parse");
        let odt = parser
            .parse_reader(std::io::Cursor::new(odt_zip))
            .expect("odt parse");

        let docx_summary = IrSummary::from_store(&docx.store);
        let odt_summary = IrSummary::from_store(&odt.store);

        assert!(docx_summary.run_texts.contains("Hello"));
        assert!(odt_summary.run_texts.contains("Hello"));
        assert!(docx_summary.run_texts.contains("Cell"));
        assert!(odt_summary.run_texts.contains("Cell"));

        assert_eq!(
            docx_summary.count(NodeType::Table),
            odt_summary.count(NodeType::Table)
        );
        assert_eq!(
            docx_summary.count(NodeType::Comment),
            odt_summary.count(NodeType::Comment)
        );
        assert_eq!(
            docx_summary.count(NodeType::Footnote),
            odt_summary.count(NodeType::Footnote)
        );
        assert_eq!(
            docx_summary.count(NodeType::Revision),
            odt_summary.count(NodeType::Revision)
        );
        assert_eq!(
            docx_summary.count(NodeType::ExternalReference),
            odt_summary.count(NodeType::ExternalReference)
        );

        std::fs::remove_file(docx_path).ok();
    }

    fn create_pptx_with_media(slide_xml: &str, slide_rels: &str) -> std::path::PathBuf {
        let mut path = std::env::temp_dir();
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_nanos();
        path.push(format!("docir_pptx_{nanos}.pptx"));

        let file = File::create(&path).expect("create temp pptx");
        let mut zip = zip::ZipWriter::new(file);
        let options = FileOptions::<()>::default();

        let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="wav" ContentType="audio/wav"/>
          <Override PartName="/ppt/presentation.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
          <Override PartName="/ppt/slides/slide1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
        </Types>"#;

        let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="ppt/presentation.xml"/>
        </Relationships>"#;

        let presentation = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId r:id="rId1"/>
          </p:sldIdLst>
        </p:presentation>"#;

        let pres_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
            Target="slides/slide1.xml"/>
        </Relationships>"#;

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.trim().as_bytes()).unwrap();
        zip.add_directory("_rels/", options).unwrap();
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(rels.trim().as_bytes()).unwrap();
        zip.add_directory("ppt/", options).unwrap();
        zip.add_directory("ppt/_rels/", options).unwrap();
        zip.start_file("ppt/presentation.xml", options).unwrap();
        zip.write_all(presentation.trim().as_bytes()).unwrap();
        zip.start_file("ppt/_rels/presentation.xml.rels", options)
            .unwrap();
        zip.write_all(pres_rels.trim().as_bytes()).unwrap();
        zip.add_directory("ppt/slides/", options).unwrap();
        zip.add_directory("ppt/slides/_rels/", options).unwrap();
        zip.start_file("ppt/slides/slide1.xml", options).unwrap();
        zip.write_all(slide_xml.trim().as_bytes()).unwrap();
        zip.start_file("ppt/slides/_rels/slide1.xml.rels", options)
            .unwrap();
        zip.write_all(slide_rels.trim().as_bytes()).unwrap();
        zip.add_directory("ppt/media/", options).unwrap();
        zip.start_file("ppt/media/audio1.wav", options).unwrap();
        zip.write_all(b"RIFFDATA").unwrap();
        zip.finish().unwrap();

        path
    }

    #[test]
    fn test_content_type_inventory_diagnostics() {
        let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
          </w:body>
        </w:document>"#;
        let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/extra.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
        </Types>"#;
        let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;
        let path = create_docx_with_relationships(
            body,
            rels,
            content_types,
            &[("word/extra.xml", "<root/>")],
        );

        let parser = OoxmlParser::new();
        let parsed = parser.parse_file(&path).expect("parse docx");
        let doc = parsed.document().expect("doc");
        assert!(!doc.diagnostics.is_empty());
        let mut found = false;
        for diag_id in &doc.diagnostics {
            if let Some(IRNode::Diagnostics(diag)) = parsed.store.get(*diag_id) {
                if diag.entries.iter().any(|e| {
                    e.code == "CONTENT_TYPE_INVENTORY" && e.message.contains("word/extra.xml")
                }) {
                    found = true;
                }
            }
        }
        assert!(found);
        std::fs::remove_file(path).ok();
    }

    #[test]
    fn test_content_type_unknown_diagnostics() {
        let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
          </w:body>
        </w:document>"#;
        let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>"#;
        let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;
        let path = create_docx_with_relationships(
            body,
            rels,
            content_types,
            &[("word/unknown.bin", "BIN")],
        );

        let parser = OoxmlParser::new();
        let parsed = parser.parse_file(&path).expect("parse docx");
        let mut found = false;
        for node in parsed.store.values() {
            if let IRNode::Diagnostics(diag) = node {
                if diag.entries.iter().any(|e| {
                    e.code == "CONTENT_TYPE_UNKNOWN" && e.message.contains("word/unknown.bin")
                }) {
                    found = true;
                }
            }
        }
        assert!(found);
        std::fs::remove_file(path).ok();
    }

    #[test]
    fn test_document_parser_enforces_max_input_size() {
        let mut config = ParserConfig::default();
        config.max_input_size = 32;
        let parser = DocumentParser::with_config(config);
        let data = vec![b'A'; 128];
        let err = parser
            .parse_reader(std::io::Cursor::new(data))
            .expect_err("expected size limit error");
        assert!(matches!(err, ParseError::ResourceLimit(_)));
    }
}
