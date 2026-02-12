//! ODF (OpenDocument) parsing support.

use crate::error::ParseError;
use crate::input::{enforce_input_size, parse_from_bytes, parse_from_file};
use crate::parser::{ParsedDocument, ParserConfig};
use crate::text_utils::parse_text_alignment;
use crate::zip_handler::SecureZipReader;
use aes::Aes128;
use aes::Aes256;
use cbc::cipher::{block_padding::Pkcs7, BlockDecryptMut, KeyIvInit};
use cbc::Decryptor;
use docir_core::ir::{
    BookmarkEnd, BookmarkStart, Cell, CellFormula, CellValue, ChartData, Comment, CommentReference,
    ConditionalFormat, ConditionalRule, DataValidation, DefinedName, DiagnosticEntry,
    DiagnosticSeverity, Diagnostics, Document, DocumentMetadata, Endnote, ExtensionPart,
    ExtensionPartKind, Field, FieldInstruction, FieldKind, Footer, Footnote, Header, IRNode,
    MediaAsset, MediaType, MergedCellRange, NumberingInfo, Paragraph, ParagraphProperties,
    PivotCache, PivotCacheRecords, PivotTable, Revision, RevisionType, Run, Section, Shape,
    ShapeText, ShapeTextParagraph, ShapeTextRun, ShapeTransform, ShapeType, Slide, SlideAnimation,
    SlideTransition, Style, StyleSet, StyleType, Table, TableCell, TableCellProperties, TableRow,
    Worksheet, WorksheetDrawing,
};
use docir_core::security::{MacroModule, MacroModuleType, MacroProject};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use pbkdf2::pbkdf2_hmac;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use sha1::Sha1;
use std::collections::HashMap;
use std::io::{Read, Seek};
use std::path::Path;
use std::sync::Arc;
use std::thread;

mod limits;
mod manifest;
mod presentation;
mod security;
mod spreadsheet;
mod text;

use self::security::{
    scan_embedded_objects, scan_external_links, scan_odf_advanced_features, scan_odf_filters,
    scan_odf_formula_security, scan_odf_objects, scan_odf_protection, OdfFormulaScan,
};

use limits::{OdfAtomicLimits, OdfLimitCounter, OdfLimits};
use manifest::{
    encrypted_manifest_entries, format_odf_encryption_metadata, is_manifest_entry_encrypted,
    parse_manifest, OdfEncryptionData, OdfManifestEntry,
};

type OdfReader<'a> = Reader<std::io::Cursor<&'a [u8]>>;

/// ODF parser (ODT/ODS/ODP).
pub struct OdfParser {
    config: ParserConfig,
}

#[derive(Debug, Default)]
struct OdfContentResult {
    content: Vec<NodeId>,
    comments: Vec<NodeId>,
    footnotes: Vec<NodeId>,
    endnotes: Vec<NodeId>,
    pivot_caches: Vec<NodeId>,
}

impl OdfParser {
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
        let mut zip = SecureZipReader::new(reader, self.config.zip_config.clone())?;

        let mimetype = zip
            .read_file_string("mimetype")
            .map(|s| s.trim().to_string())
            .map_err(|_| ParseError::UnsupportedFormat("Missing ODF mimetype".to_string()))?;

        let format = detect_odf_format(&mimetype).ok_or_else(|| {
            ParseError::UnsupportedFormat(format!("Unsupported ODF mimetype: {mimetype}"))
        })?;

        let manifest_entries = if zip.contains("META-INF/manifest.xml") {
            let manifest_xml = zip.read_file_string("META-INF/manifest.xml")?;
            parse_manifest(&manifest_xml)?
        } else {
            Vec::new()
        };

        if !zip.contains("content.xml") {
            return Err(ParseError::MissingPart("content.xml".to_string()));
        }

        let mut store = IrStore::new();
        let mut doc = Document::new(format);
        let mut diagnostics = Diagnostics::new();

        load_meta(&mut zip, &mut store, &mut doc)?;
        let content_state = handle_content_xml(
            &self.config,
            &mut zip,
            format,
            &manifest_entries,
            &mut store,
            &mut doc,
            &mut diagnostics,
        )?;
        let content_xml = content_state.content_xml;
        let content_bytes = content_state.content_bytes;
        let fast_mode = content_state.fast_mode;
        let content_size = content_state.content_size;

        let mut styles_xml: Option<String> = None;
        if zip.contains("styles.xml") {
            let xml = zip.read_file_string("styles.xml")?;
            if let Some(styles) = parse_styles(&xml) {
                let style_id = styles.id;
                store.insert(IRNode::StyleSet(styles));
                doc.styles = Some(style_id);
            }
            styles_xml = Some(xml);
        }

        let settings_xml = if zip.contains("settings.xml") {
            Some(zip.read_file_string("settings.xml")?)
        } else {
            None
        };

        let signatures_xml = if zip.contains("META-INF/documentsignatures.xml") {
            Some(zip.read_file_string("META-INF/documentsignatures.xml")?)
        } else {
            None
        };

        if fast_mode {
            let size = content_size.unwrap_or(0);
            diagnostics.entries.push(DiagnosticEntry {
                severity: DiagnosticSeverity::Info,
                code: "ODF_FAST_MODE".to_string(),
                message: format!(
                    "ODF fast mode enabled (content.xml: {} bytes, threshold: {} bytes, sample_rows: {}, sample_cols: {})",
                    size,
                    self.config.odf_fast_threshold_bytes,
                    self.config.odf_fast_sample_rows,
                    self.config.odf_fast_sample_cols
                ),
                path: Some("content.xml".to_string()),
            });
            if content_xml.is_none() {
                diagnostics.entries.push(DiagnosticEntry {
                    severity: DiagnosticSeverity::Warning,
                    code: "ODF_FAST_SKIP_SCAN".to_string(),
                    message:
                        "Fast mode skipped content.xml security scans to reduce processing time"
                            .to_string(),
                    path: Some("content.xml".to_string()),
                });
            }
        }
        let mut manifest_index: HashMap<String, Option<String>> = HashMap::new();
        for entry in &manifest_entries {
            let path = entry.path.clone();
            let media_type = entry.media_type.clone();
            manifest_index.insert(path.clone(), media_type.clone());
            diagnostics.entries.push(DiagnosticEntry {
                severity: DiagnosticSeverity::Info,
                code: "ODF_PART".to_string(),
                message: format!(
                    "ODF part: {} (media-type: {})",
                    path,
                    media_type.clone().unwrap_or_else(|| "(none)".to_string())
                ),
                path: Some(path),
            });
        }

        let mut file_names: Vec<String> = zip.file_names().map(|name| name.to_string()).collect();
        file_names.sort();
        for path in &file_names {
            if path == "mimetype" {
                continue;
            }
            let media_type = manifest_index.get(path.as_str()).cloned().unwrap_or(None);
            let size_bytes = zip.file_size(&path).unwrap_or(0);
            let mut part =
                ExtensionPart::new(path.to_string(), size_bytes, ExtensionPartKind::Unknown);
            part.content_type = media_type.clone();
            let part_id = part.id;
            store.insert(IRNode::ExtensionPart(part));
            doc.shared_parts.push(part_id);

            if let Some(media) = media_type.as_deref() {
                if let Some(asset) = build_media_asset(path, media, size_bytes) {
                    let asset_id = asset.id;
                    store.insert(IRNode::MediaAsset(asset));
                    doc.shared_parts.push(asset_id);
                }
            }
        }

        if let Some(xml) = styles_xml.as_deref() {
            let masters = parse_master_pages(xml);
            for name in masters {
                diagnostics.entries.push(DiagnosticEntry {
                    severity: DiagnosticSeverity::Info,
                    code: "ODF_MASTER_PAGE".to_string(),
                    message: format!("ODF master page detected: {}", name),
                    path: Some("styles.xml".to_string()),
                });
            }
            let layouts = parse_page_layouts(xml);
            for name in layouts {
                diagnostics.entries.push(DiagnosticEntry {
                    severity: DiagnosticSeverity::Info,
                    code: "ODF_PAGE_LAYOUT".to_string(),
                    message: format!("ODF page layout detected: {}", name),
                    path: Some("styles.xml".to_string()),
                });
            }
            let (headers, footers) = parse_odf_headers_footers(xml, &mut store, &self.config)?;
            for header_id in headers {
                doc.shared_parts.push(header_id);
            }
            for footer_id in footers {
                doc.shared_parts.push(footer_id);
            }
        }

        if !fast_mode {
            if let Some(xml) = content_xml.as_deref() {
                if let Some(mut styles) = parse_styles(xml) {
                    if let Some(doc_styles_id) = doc.styles {
                        if let Some(IRNode::StyleSet(existing)) = store.get_mut(doc_styles_id) {
                            merge_styles(existing, &mut styles);
                        }
                    } else {
                        let style_id = styles.id;
                        store.insert(IRNode::StyleSet(styles));
                        doc.styles = Some(style_id);
                    }
                }
                let masters = parse_master_pages(xml);
                for name in masters {
                    diagnostics.entries.push(DiagnosticEntry {
                        severity: DiagnosticSeverity::Info,
                        code: "ODF_MASTER_PAGE".to_string(),
                        message: format!("ODF master page detected: {}", name),
                        path: Some("content.xml".to_string()),
                    });
                }
                let layouts = parse_page_layouts(xml);
                for name in layouts {
                    diagnostics.entries.push(DiagnosticEntry {
                        severity: DiagnosticSeverity::Info,
                        code: "ODF_PAGE_LAYOUT".to_string(),
                        message: format!("ODF page layout detected: {}", name),
                        path: Some("content.xml".to_string()),
                    });
                }
            }
        }

        if fast_mode && format == DocumentFormat::OdfSpreadsheet {
            if let Some(bytes) = content_bytes.as_deref() {
                let chunks = spreadsheet::extract_spreadsheet_table_chunks(bytes);
                for (idx, chunk) in chunks.iter().enumerate() {
                    let sheet_name =
                        spreadsheet::table_name_from_chunk(&chunk.bytes, (idx + 1) as u32);
                    let path = format!(
                        "content.xml#sheet:{}@{}-{}",
                        sheet_name, chunk.start, chunk.end
                    );
                    let mut part = ExtensionPart::new(
                        path,
                        (chunk.end.saturating_sub(chunk.start) + 1) as u64,
                        ExtensionPartKind::VendorSpecific,
                    );
                    part.content_type =
                        Some("application/vnd.docir.odf.lazy-sheet+xml".to_string());
                    part.span = Some(SourceSpan::new("content.xml"));
                    let part_id = part.id;
                    store.insert(IRNode::ExtensionPart(part));
                    doc.shared_parts.push(part_id);
                    diagnostics.entries.push(DiagnosticEntry {
                        severity: DiagnosticSeverity::Info,
                        code: "ODF_LAZY_SHEET".to_string(),
                        message: format!(
                            "Lazy sheet range stored for {} ({}-{})",
                            sheet_name, chunk.start, chunk.end
                        ),
                        path: Some("content.xml".to_string()),
                    });
                }
            }
        }

        let mut security = doc.security.clone();
        let mut macro_project = build_odf_macro_project(
            &manifest_entries,
            &content_xml,
            &styles_xml,
            &settings_xml,
            &file_names,
            &mut store,
        );
        if let Some(project) = macro_project.take() {
            let project_id = project.id;
            store.insert(IRNode::MacroProject(project));
            security.macro_project = Some(project_id);
        }

        let mut formula_scan = OdfFormulaScan::default();
        if let Some(xml) = content_xml.as_deref() {
            formula_scan = scan_odf_formula_security(xml);
            diagnostics
                .entries
                .extend(formula_scan.diagnostics.drain(..));
            diagnostics.entries.extend(scan_odf_protection(xml));
            diagnostics.entries.extend(scan_odf_advanced_features(xml));
        }

        let mut external_refs = Vec::new();
        if let Some(xml) = content_xml.as_deref() {
            external_refs.extend(scan_external_links(xml, "content.xml"));
        }
        if let Some(xml) = styles_xml.as_deref() {
            external_refs.extend(scan_external_links(xml, "styles.xml"));
        }
        if let Some(xml) = settings_xml.as_deref() {
            external_refs.extend(scan_external_links(xml, "settings.xml"));
        }
        external_refs.extend(formula_scan.external_refs.drain(..));

        let mut ole_objects = Vec::new();
        if let Some(xml) = content_xml.as_deref() {
            let (oles, ole_links) = scan_odf_objects(xml);
            ole_objects.extend(oles);
            external_refs.extend(ole_links);
        }
        ole_objects.extend(scan_embedded_objects(&file_names, &mut zip));

        for ext in external_refs {
            let id = ext.id;
            store.insert(IRNode::ExternalReference(ext));
            security.external_refs.push(id);
        }
        for ole in ole_objects {
            let id = ole.id;
            store.insert(IRNode::OleObject(ole));
            security.ole_objects.push(id);
        }
        security
            .dde_fields
            .extend(formula_scan.dde_fields.drain(..));
        doc.security = security;

        if let Some(sig_xml) = signatures_xml.as_deref() {
            let sigs = parse_odf_signatures(sig_xml);
            for sig in sigs {
                let sig_id = sig.id;
                store.insert(IRNode::DigitalSignature(sig));
                doc.shared_parts.push(sig_id);
            }
        }

        let encrypted_entries = encrypted_manifest_entries(&manifest_entries);
        for entry in &manifest_entries {
            if let Some(message) = format_odf_encryption_metadata(entry) {
                diagnostics.entries.push(DiagnosticEntry {
                    severity: DiagnosticSeverity::Info,
                    code: "ODF_ENCRYPTION_META".to_string(),
                    message,
                    path: Some("META-INF/manifest.xml".to_string()),
                });
            }
        }
        if !encrypted_entries.is_empty() {
            diagnostics.entries.push(DiagnosticEntry {
                severity: DiagnosticSeverity::Warning,
                code: "ODF_ENCRYPTION".to_string(),
                message: "ODF encrypted content detected in manifest".to_string(),
                path: Some("META-INF/manifest.xml".to_string()),
            });
            for path in encrypted_entries {
                diagnostics.entries.push(DiagnosticEntry {
                    severity: DiagnosticSeverity::Warning,
                    code: "ODF_ENCRYPTED_PART".to_string(),
                    message: format!("Encrypted ODF part not decrypted: {}", path),
                    path: Some("META-INF/manifest.xml".to_string()),
                });
            }
        }

        if !diagnostics.entries.is_empty() {
            let diag_id = diagnostics.id;
            store.insert(IRNode::Diagnostics(diagnostics));
            doc.diagnostics.push(diag_id);
        }

        if let Some(xml) = content_xml.as_deref() {
            for filter in scan_odf_filters(xml) {
                let mut diag = Diagnostics::new();
                diag.entries.push(DiagnosticEntry {
                    severity: DiagnosticSeverity::Info,
                    code: "ODF_FILTER".to_string(),
                    message: format!("ODF filter detected: {}", filter),
                    path: Some("content.xml".to_string()),
                });
                let diag_id = diag.id;
                store.insert(IRNode::Diagnostics(diag));
                doc.diagnostics.push(diag_id);
            }
        }

        if format == DocumentFormat::OdfSpreadsheet {
            if let Some(bytes) = content_bytes.as_deref() {
                let defined_names = parse_ods_named_ranges(bytes);
                for name in defined_names {
                    let id = name.id;
                    store.insert(IRNode::DefinedName(name));
                    doc.defined_names.push(id);
                }
            }
        }

        let doc_id = doc.id;
        store.insert(IRNode::Document(doc));

        Ok(ParsedDocument {
            root_id: doc_id,
            format,
            store,
            metrics: None,
        })
    }
}

struct ContentState {
    content_xml: Option<String>,
    content_bytes: Option<Vec<u8>>,
    fast_mode: bool,
    content_size: Option<u64>,
}

fn load_meta<R: Read + Seek>(
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
    doc: &mut Document,
) -> Result<(), ParseError> {
    if zip.contains("meta.xml") {
        let meta_xml = zip.read_file_string("meta.xml")?;
        if let Some(meta) = parse_meta(&meta_xml) {
            let meta_id = meta.id;
            store.insert(IRNode::Metadata(meta));
            doc.metadata = Some(meta_id);
        }
    }
    Ok(())
}

fn handle_content_xml<R: Read + Seek>(
    config: &ParserConfig,
    zip: &mut SecureZipReader<R>,
    format: DocumentFormat,
    manifest_entries: &[OdfManifestEntry],
    store: &mut IrStore,
    doc: &mut Document,
    diagnostics: &mut Diagnostics,
) -> Result<ContentState, ParseError> {
    let mut content_xml: Option<String> = None;
    let mut content_bytes: Option<Vec<u8>> = None;
    let mut fast_mode = false;
    let mut content_size = None;
    let content_entry = manifest_entries
        .iter()
        .find(|entry| entry.path == "content.xml");
    let content_encrypted = content_entry
        .map(is_manifest_entry_encrypted)
        .unwrap_or(false);
    if zip.contains("content.xml") {
        let size = zip.file_size("content.xml")?;
        content_size = Some(size);
        if let Some(max_bytes) = config.odf_max_bytes {
            if size > max_bytes {
                return Err(ParseError::ResourceLimit(format!(
                    "ODF content.xml too large: {} bytes (max: {} bytes)",
                    size, max_bytes
                )));
            }
        }
        if format == DocumentFormat::OdfSpreadsheet
            && (config.odf_force_fast || size >= config.odf_fast_threshold_bytes)
        {
            fast_mode = true;
        }

        let xml_bytes = if content_encrypted {
            let password = config.odf_password.as_deref();
            let encryption = content_entry.and_then(|entry| entry.encryption.as_ref());
            if let (Some(password), Some(encryption)) = (password, encryption) {
                match decrypt_odf_part(zip.read_file("content.xml")?, encryption, password) {
                    Ok(bytes) => {
                        diagnostics.entries.push(DiagnosticEntry {
                            severity: DiagnosticSeverity::Info,
                            code: "ODF_DECRYPT_OK".to_string(),
                            message: "ODF encrypted content.xml decrypted successfully".to_string(),
                            path: Some("content.xml".to_string()),
                        });
                        bytes
                    }
                    Err(message) => {
                        return Err(ParseError::InvalidFormat(format!(
                            "ODF decryption failed: {}",
                            message
                        )));
                    }
                }
            } else {
                diagnostics.entries.push(DiagnosticEntry {
                    severity: DiagnosticSeverity::Warning,
                    code: "ODF_DECRYPT_SKIPPED".to_string(),
                    message:
                        "ODF content.xml is encrypted but no password or encryption data is available"
                            .to_string(),
                    path: Some("content.xml".to_string()),
                });
                Vec::new()
            }
        } else {
            zip.read_file("content.xml")?
        };
        if xml_bytes.is_empty() {
            content_bytes = Some(xml_bytes);
        } else {
            if !fast_mode {
                content_xml = Some(String::from_utf8_lossy(&xml_bytes).to_string());
            }
            let use_parallel = format == DocumentFormat::OdfSpreadsheet
                && config.odf_parallel_sheets
                && !fast_mode;
            let content_result = if use_parallel {
                let limits = Arc::new(OdfAtomicLimits::new(config, fast_mode));
                spreadsheet::parse_content_spreadsheet_parallel(&xml_bytes, store, &limits, config)?
            } else {
                let limits = OdfLimits::new(config, fast_mode);
                parse_content(&xml_bytes, format, store, &limits)?
            };
            doc.content.extend(content_result.content);
            doc.comments.extend(content_result.comments);
            doc.footnotes.extend(content_result.footnotes);
            doc.endnotes.extend(content_result.endnotes);
            doc.pivot_caches.extend(content_result.pivot_caches);
            content_bytes = Some(xml_bytes);
        }
    }

    Ok(ContentState {
        content_xml,
        content_bytes,
        fast_mode,
        content_size,
    })
}

fn detect_odf_format(mimetype: &str) -> Option<DocumentFormat> {
    let lower = mimetype.to_ascii_lowercase();
    if lower.contains("opendocument.text") || lower.contains("vnd.sun.xml.writer") {
        Some(DocumentFormat::OdfText)
    } else if lower.contains("opendocument.spreadsheet") || lower.contains("vnd.sun.xml.calc") {
        Some(DocumentFormat::OdfSpreadsheet)
    } else if lower.contains("opendocument.presentation") || lower.contains("vnd.sun.xml.impress") {
        Some(DocumentFormat::OdfPresentation)
    } else {
        None
    }
}

fn decrypt_odf_part(
    encrypted: Vec<u8>,
    encryption: &OdfEncryptionData,
    password: &str,
) -> Result<Vec<u8>, String> {
    let algorithm = encryption
        .algorithm_name
        .as_deref()
        .unwrap_or("http://www.w3.org/2001/04/xmlenc#aes256-cbc");
    let salt = encryption
        .salt
        .as_ref()
        .ok_or_else(|| "Missing encryption salt".to_string())?;
    let iv = encryption
        .init_vector
        .as_ref()
        .ok_or_else(|| "Missing encryption IV".to_string())?;
    let iterations = encryption.iteration_count.unwrap_or(100_000);
    let key_bits = encryption
        .key_size
        .or_else(|| {
            if algorithm.contains("aes256") {
                Some(256)
            } else if algorithm.contains("aes128") {
                Some(128)
            } else {
                None
            }
        })
        .ok_or_else(|| "Unsupported encryption algorithm".to_string())?;
    let key_len = (key_bits / 8) as usize;
    if iv.len() != 16 {
        return Err(format!("Unsupported IV length: {}", iv.len()));
    }

    let mut key = vec![0u8; key_len];
    pbkdf2_hmac::<Sha1>(password.as_bytes(), salt, iterations, &mut key);

    let mut buffer = encrypted;
    if key_len == 32 {
        let decryptor = Decryptor::<Aes256>::new_from_slices(&key, iv)
            .map_err(|_| "Invalid AES-256 key or IV".to_string())?;
        let decrypted = decryptor
            .decrypt_padded_mut::<Pkcs7>(&mut buffer)
            .map_err(|_| "Invalid AES-256 padding".to_string())?;
        Ok(decrypted.to_vec())
    } else if key_len == 16 {
        let decryptor = Decryptor::<Aes128>::new_from_slices(&key, iv)
            .map_err(|_| "Invalid AES-128 key or IV".to_string())?;
        let decrypted = decryptor
            .decrypt_padded_mut::<Pkcs7>(&mut buffer)
            .map_err(|_| "Invalid AES-128 padding".to_string())?;
        Ok(decrypted.to_vec())
    } else {
        Err(format!("Unsupported key length: {}", key_len))
    }
}

fn parse_meta(xml: &str) -> Option<DocumentMetadata> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut meta = DocumentMetadata::new();
    let mut current = None;

    #[derive(Clone, Copy)]
    enum MetaField {
        Title,
        Subject,
        Creator,
        Keywords,
        Description,
        Created,
        Modified,
    }

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                current = match e.name().as_ref() {
                    b"dc:title" => Some(MetaField::Title),
                    b"dc:subject" => Some(MetaField::Subject),
                    b"dc:creator" => Some(MetaField::Creator),
                    b"meta:keyword" => Some(MetaField::Keywords),
                    b"dc:description" => Some(MetaField::Description),
                    b"meta:creation-date" => Some(MetaField::Created),
                    b"dc:date" => Some(MetaField::Modified),
                    _ => None,
                };
            }
            Ok(Event::Text(e)) => {
                if let Some(field) = current {
                    let value = e.unescape().unwrap_or_default().to_string();
                    match field {
                        MetaField::Title => meta.title = Some(value),
                        MetaField::Subject => meta.subject = Some(value),
                        MetaField::Creator => meta.creator = Some(value),
                        MetaField::Keywords => meta.keywords = Some(value),
                        MetaField::Description => meta.description = Some(value),
                        MetaField::Created => meta.created = Some(value),
                        MetaField::Modified => meta.modified = Some(value),
                    }
                }
            }
            Ok(Event::End(_)) => {
                current = None;
            }
            Ok(Event::Eof) => break,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }

    let has_any = meta.title.is_some()
        || meta.subject.is_some()
        || meta.creator.is_some()
        || meta.keywords.is_some()
        || meta.description.is_some()
        || meta.created.is_some()
        || meta.modified.is_some();

    if has_any {
        Some(meta)
    } else {
        None
    }
}

fn parse_content(
    xml: &[u8],
    format: DocumentFormat,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<OdfContentResult, ParseError> {
    match format {
        DocumentFormat::OdfText => text::parse_content_text(xml, store, limits),
        DocumentFormat::OdfSpreadsheet => {
            spreadsheet::parse_content_spreadsheet(xml, store, limits)
        }
        DocumentFormat::OdfPresentation => {
            presentation::parse_content_presentation(xml, store, limits)
        }
        _ => Ok(OdfContentResult::default()),
    }
}

fn attr_value(e: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == name {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

fn build_odf_macro_project(
    manifest_entries: &[OdfManifestEntry],
    content_xml: &Option<String>,
    styles_xml: &Option<String>,
    settings_xml: &Option<String>,
    file_names: &[String],
    store: &mut IrStore,
) -> Option<MacroProject> {
    let mut module_paths = Vec::new();
    for entry in manifest_entries {
        if let Some(media) = entry.media_type.as_deref() {
            if media.contains("script") || media.contains("basic") {
                module_paths.push(entry.path.clone());
            }
        }
    }
    for name in file_names {
        if name.starts_with("Scripts/") || name.starts_with("Basic/") || name.ends_with(".bas") {
            module_paths.push(name.clone());
        }
    }

    if let Some(xml) = content_xml.as_deref() {
        module_paths.extend(scan_script_links(xml));
    }
    if let Some(xml) = styles_xml.as_deref() {
        module_paths.extend(scan_script_links(xml));
    }
    if let Some(xml) = settings_xml.as_deref() {
        module_paths.extend(scan_script_links(xml));
    }

    module_paths.sort();
    module_paths.dedup();

    if module_paths.is_empty() {
        return None;
    }

    let mut project = MacroProject::new();
    project.name = Some("ODF Scripts".to_string());

    for path in module_paths {
        let module = MacroModule::new(path.clone(), MacroModuleType::Standard);
        let module_id = module.id;
        store.insert(IRNode::MacroModule(module));
        project.modules.push(module_id);
    }

    Some(project)
}

fn scan_script_links(xml: &str) -> Vec<String> {
    let mut links = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"script:script" {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        links.push(href);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    links
}

fn strip_odf_formula_prefix(formula: &str) -> &str {
    if let Some(stripped) = formula.strip_prefix("of:=") {
        stripped
    } else if let Some(stripped) = formula.strip_prefix("of:") {
        stripped
    } else {
        formula
    }
}

fn parse_odf_signatures(xml: &str) -> Vec<docir_core::ir::DigitalSignature> {
    let mut sigs = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut current: Option<docir_core::ir::DigitalSignature> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"ds:Signature" => current = Some(docir_core::ir::DigitalSignature::new()),
                b"ds:SignatureMethod" => {
                    if let Some(sig) = current.as_mut() {
                        sig.signature_method = attr_value(&e, b"Algorithm");
                    }
                }
                b"ds:DigestMethod" => {
                    if let Some(sig) = current.as_mut() {
                        if let Some(alg) = attr_value(&e, b"Algorithm") {
                            sig.digest_methods.push(alg);
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if let Some(sig) = current.as_mut() {
                    let text = e.unescape().unwrap_or_default().to_string();
                    if sig.signer.is_none() && text.contains("CN=") {
                        sig.signer = Some(text);
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"ds:Signature" {
                    if let Some(sig) = current.take() {
                        sigs.push(sig);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    sigs
}

fn parse_draw_page(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    slide_no: u32,
    store: &mut IrStore,
) -> Result<Slide, ParseError> {
    let mut slide = Slide::new(slide_no);
    slide.name = attr_value(start, b"draw:name");
    slide.master_id = attr_value(start, b"draw:master-page-name")
        .or_else(|| attr_value(start, b"presentation:master-page-name"));
    slide.layout_id = attr_value(start, b"presentation:page-layout-name")
        .or_else(|| attr_value(start, b"draw:style-name"));
    slide.transition = parse_odp_transition(start);
    let mut buf = Vec::new();
    let mut notes_text: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"draw:frame" => {
                    if let Some(shape_id) = parse_draw_frame_presentation(reader, &e, store)? {
                        slide.shapes.push(shape_id);
                    }
                }
                b"draw:custom-shape" => {
                    if let Some(shape_id) = parse_custom_shape_presentation(reader, &e, store)? {
                        slide.shapes.push(shape_id);
                    }
                }
                b"presentation:notes" => {
                    notes_text = parse_notes(reader)?;
                }
                name if name.starts_with(b"anim:") => {
                    if let Some(anim) = parse_odf_animation(&e) {
                        slide.animations.push(anim);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"draw:frame" => {
                    if let Some(shape_id) = parse_draw_frame_presentation(reader, &e, store)? {
                        slide.shapes.push(shape_id);
                    }
                }
                b"draw:custom-shape" => {
                    if let Some(shape_id) = parse_custom_shape_presentation(reader, &e, store)? {
                        slide.shapes.push(shape_id);
                    }
                }
                name if name.starts_with(b"anim:") => {
                    if let Some(anim) = parse_odf_animation(&e) {
                        slide.animations.push(anim);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"draw:page" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    slide.notes = notes_text;
    Ok(slide)
}

fn parse_odp_transition(start: &BytesStart<'_>) -> Option<SlideTransition> {
    let transition_type = attr_value(start, b"presentation:transition-type")
        .or_else(|| attr_value(start, b"draw:transition-type"));
    let speed = attr_value(start, b"presentation:transition-speed");
    let duration_ms =
        attr_value(start, b"presentation:transition-duration").and_then(|v| v.parse::<u32>().ok());
    let advance_after_ms =
        attr_value(start, b"presentation:duration").and_then(|v| v.parse::<u32>().ok());
    let advance_on_click = attr_value(start, b"presentation:animation").map(|v| v == "click");
    if transition_type.is_some() || speed.is_some() || duration_ms.is_some() {
        Some(SlideTransition {
            transition_type,
            speed,
            advance_on_click,
            advance_after_ms,
            duration_ms,
        })
    } else {
        None
    }
}

fn parse_odf_chart(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
) -> Result<ChartData, ParseError> {
    let mut chart = ChartData::new();
    chart.chart_type = attr_value(start, b"chart:class");
    chart.span = Some(SourceSpan::new("content.xml"));
    let mut buf = Vec::new();
    let mut in_title = false;
    let mut title_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"chart:title" => {
                    in_title = true;
                }
                b"text:p" if in_title => {
                    let text = parse_text_element(reader, b"text:p")?;
                    if !title_text.is_empty() && !text.is_empty() {
                        title_text.push(' ');
                    }
                    title_text.push_str(&text);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"chart:title" {
                    in_title = false;
                }
                if e.name().as_ref() == b"chart:chart" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if !title_text.is_empty() {
        chart.title = Some(title_text);
    }
    Ok(chart)
}

fn parse_odf_animation(start: &BytesStart<'_>) -> Option<SlideAnimation> {
    let name = String::from_utf8_lossy(start.name().as_ref()).to_string();
    let mut anim = SlideAnimation {
        animation_type: name,
        target: attr_value(start, b"anim:targetElement"),
        duration_ms: attr_value(start, b"smil:dur").and_then(|v| parse_duration_ms(&v)),
        preset_id: attr_value(start, b"presentation:preset-id"),
        preset_class: attr_value(start, b"presentation:preset-class"),
        media_asset: None,
    };
    if anim.target.is_none() {
        anim.target = attr_value(start, b"smil:targetElement");
    }
    Some(anim)
}

fn parse_duration_ms(value: &str) -> Option<u32> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return None;
    }
    if let Some(stripped) = trimmed.strip_suffix("ms") {
        return stripped.parse::<u32>().ok();
    }
    if let Some(stripped) = trimmed.strip_suffix('s') {
        return stripped
            .parse::<f32>()
            .ok()
            .map(|v| (v * 1000.0).round() as u32);
    }
    if trimmed.starts_with("PT") && trimmed.ends_with('S') {
        let inner = trimmed.trim_start_matches("PT").trim_end_matches('S');
        return inner
            .parse::<f32>()
            .ok()
            .map(|v| (v * 1000.0).round() as u32);
    }
    None
}

fn build_media_asset(path: &str, media: &str, size_bytes: u64) -> Option<MediaAsset> {
    let media_type = classify_media_type(path, media)?;
    let mut asset = MediaAsset::new(path.to_string(), media_type, size_bytes);
    asset.content_type = Some(media.to_string());
    asset.span = Some(SourceSpan::new("META-INF/manifest.xml"));
    Some(asset)
}

fn classify_media_type(path: &str, media: &str) -> Option<MediaType> {
    let lower_media = media.to_ascii_lowercase();
    if lower_media.starts_with("image/") {
        return Some(MediaType::Image);
    }
    if lower_media.starts_with("audio/") {
        return Some(MediaType::Audio);
    }
    if lower_media.starts_with("video/") {
        return Some(MediaType::Video);
    }
    if lower_media.starts_with("application/") {
        let lower_path = path.to_ascii_lowercase();
        if lower_path.ends_with(".ogg") || lower_path.ends_with(".oga") {
            return Some(MediaType::Audio);
        }
        if lower_path.ends_with(".ogv") {
            return Some(MediaType::Video);
        }
    }
    None
}

fn classify_media_shape(path: &str) -> ShapeType {
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".mp3")
        || lower.ends_with(".wav")
        || lower.ends_with(".ogg")
        || lower.ends_with(".oga")
    {
        return ShapeType::Audio;
    }
    if lower.ends_with(".mp4")
        || lower.ends_with(".mov")
        || lower.ends_with(".avi")
        || lower.ends_with(".mpeg")
        || lower.ends_with(".mpg")
        || lower.ends_with(".ogv")
    {
        return ShapeType::Video;
    }
    ShapeType::Unknown
}

fn parse_ods_named_ranges(xml: &[u8]) -> Vec<DefinedName> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:named-range" => {
                    if let Some(name) = attr_value(&e, b"table:name") {
                        let value = attr_value(&e, b"table:cell-range-address")
                            .unwrap_or_else(|| String::new());
                        let mut def = DefinedName {
                            id: NodeId::new(),
                            name,
                            value,
                            local_sheet_id: None,
                            hidden: false,
                            comment: attr_value(&e, b"table:comment"),
                            span: Some(SourceSpan::new("content.xml")),
                        };
                        if let Some(hidden) = attr_value(&e, b"table:hidden") {
                            def.hidden = hidden == "true";
                        }
                        out.push(def);
                    }
                }
                b"table:named-expression" => {
                    if let Some(name) = attr_value(&e, b"table:name") {
                        let value =
                            attr_value(&e, b"table:expression").unwrap_or_else(|| String::new());
                        let mut def = DefinedName {
                            id: NodeId::new(),
                            name,
                            value,
                            local_sheet_id: None,
                            hidden: false,
                            comment: attr_value(&e, b"table:comment"),
                            span: Some(SourceSpan::new("content.xml")),
                        };
                        if let Some(hidden) = attr_value(&e, b"table:hidden") {
                            def.hidden = hidden == "true";
                        }
                        out.push(def);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    out
}

fn parse_frame_transform(start: &BytesStart<'_>) -> ShapeTransform {
    let mut transform = ShapeTransform::default();
    if let Some(x) = attr_value(start, b"svg:x").and_then(parse_length_emu) {
        transform.x = x;
    }
    if let Some(y) = attr_value(start, b"svg:y").and_then(parse_length_emu) {
        transform.y = y;
    }
    if let Some(width) = attr_value(start, b"svg:width").and_then(parse_length_emu_u64) {
        transform.width = width;
    }
    if let Some(height) = attr_value(start, b"svg:height").and_then(parse_length_emu_u64) {
        transform.height = height;
    }
    transform
}

fn parse_length_emu(value: String) -> Option<i64> {
    parse_length_emu_str(&value).map(|v| v.round() as i64)
}

fn parse_length_emu_u64(value: String) -> Option<u64> {
    parse_length_emu_str(&value).map(|v| v.max(0.0).round() as u64)
}

fn parse_length_emu_str(value: &str) -> Option<f64> {
    let trimmed = value.trim();
    let mut num = String::new();
    let mut unit = String::new();
    for ch in trimmed.chars() {
        if ch.is_ascii_digit() || ch == '.' || ch == '-' {
            num.push(ch);
        } else {
            unit.push(ch);
        }
    }
    let magnitude = num.parse::<f64>().ok()?;
    let emu = match unit.as_str() {
        "cm" => magnitude / 2.54 * 914_400.0,
        "mm" => magnitude / 25.4 * 914_400.0,
        "in" => magnitude * 914_400.0,
        "pt" => magnitude * 12_700.0,
        "pc" => magnitude * 152_400.0,
        "px" => magnitude * 9_525.0,
        "" => magnitude,
        _ => return None,
    };
    Some(emu)
}

fn parse_draw_frame_presentation(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let transform = parse_frame_transform(start);
    let mut shape_type = ShapeType::Picture;
    let mut media_target: Option<String> = None;
    let mut text: Option<ShapeText> = None;
    let mut name = attr_value(start, b"draw:name");
    let mut chart_id: Option<NodeId> = None;
    let mut buf = Vec::new();
    let mut has_shape = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"draw:text-box" => {
                    let paragraphs = parse_shape_text(reader, b"draw:text-box")?;
                    if !paragraphs.is_empty() {
                        text = Some(ShapeText { paragraphs });
                        shape_type = ShapeType::TextBox;
                        has_shape = true;
                    }
                }
                b"draw:image" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href);
                        shape_type = ShapeType::Picture;
                        has_shape = true;
                    }
                }
                b"chart:chart" => {
                    shape_type = ShapeType::Chart;
                    has_shape = true;
                    let chart = parse_odf_chart(reader, &e)?;
                    let id = chart.id;
                    store.insert(IRNode::ChartData(chart));
                    chart_id = Some(id);
                }
                b"draw:plugin" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                        shape_type = classify_media_shape(&href);
                        has_shape = true;
                    }
                }
                b"draw:object" | b"draw:object-ole" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                    }
                    shape_type = ShapeType::OleObject;
                    has_shape = true;
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"draw:image" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href);
                        shape_type = ShapeType::Picture;
                        has_shape = true;
                    }
                }
                b"chart:chart" => {
                    shape_type = ShapeType::Chart;
                    has_shape = true;
                    let mut chart = ChartData::new();
                    chart.chart_type = attr_value(&e, b"chart:class");
                    chart.span = Some(SourceSpan::new("content.xml"));
                    let id = chart.id;
                    store.insert(IRNode::ChartData(chart));
                    chart_id = Some(id);
                }
                b"draw:plugin" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                        shape_type = classify_media_shape(&href);
                        has_shape = true;
                    }
                }
                b"draw:object" | b"draw:object-ole" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                    }
                    shape_type = ShapeType::OleObject;
                    has_shape = true;
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"draw:frame" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if has_shape {
        let mut shape = Shape::new(shape_type);
        shape.name = name.take();
        shape.media_target = media_target;
        shape.text = text;
        shape.chart_id = chart_id;
        shape.transform = transform;
        let shape_id = shape.id;
        store.insert(IRNode::Shape(shape));
        Ok(Some(shape_id))
    } else {
        Ok(None)
    }
}

fn parse_custom_shape_presentation(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let mut name = attr_value(start, b"draw:name");
    let paragraphs = parse_shape_text(reader, b"draw:custom-shape")?;
    let mut shape = Shape::new(ShapeType::Custom);
    shape.name = name.take();
    if !paragraphs.is_empty() {
        shape.text = Some(ShapeText { paragraphs });
    }
    let shape_id = shape.id;
    store.insert(IRNode::Shape(shape));
    Ok(Some(shape_id))
}

fn parse_shape_text(
    reader: &mut OdfReader<'_>,
    end_tag: &[u8],
) -> Result<Vec<ShapeTextParagraph>, ParseError> {
    let mut buf = Vec::new();
    let mut paragraphs = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:p" => {
                    let text = parse_text_element(reader, b"text:p")?;
                    let run = ShapeTextRun {
                        text,
                        bold: None,
                        italic: None,
                        font_size: None,
                        font_family: None,
                    };
                    paragraphs.push(ShapeTextParagraph {
                        runs: vec![run],
                        alignment: None,
                    });
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
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(paragraphs)
}

fn parse_notes(reader: &mut OdfReader<'_>) -> Result<Option<String>, ParseError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"text:p" {
                    let para = parse_text_element(reader, b"text:p")?;
                    if !text.is_empty() {
                        text.push('\n');
                    }
                    text.push_str(&para);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"presentation:notes" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if text.is_empty() {
        Ok(None)
    } else {
        Ok(Some(text))
    }
}
#[derive(Debug, Clone)]
struct ValidationDef {
    validation_type: Option<String>,
    operator: Option<String>,
    allow_blank: bool,
    show_input_message: bool,
    show_error_message: bool,
    error_title: Option<String>,
    error: Option<String>,
    prompt_title: Option<String>,
    prompt: Option<String>,
    formula1: Option<String>,
    formula2: Option<String>,
}

fn parse_validation_definition(start: &BytesStart<'_>) -> Option<(String, ValidationDef)> {
    let name = attr_value(start, b"table:name")?;
    let condition = attr_value(start, b"table:condition");
    let allow_blank = attr_value(start, b"table:allow-empty-cell")
        .map(|v| v == "true")
        .unwrap_or(false);
    let show_input_message = attr_value(start, b"table:display-list")
        .map(|v| v == "true")
        .unwrap_or(false);
    let show_error_message = attr_value(start, b"table:display-list")
        .map(|v| v == "true")
        .unwrap_or(false);
    let def = ValidationDef {
        validation_type: condition.clone(),
        operator: None,
        allow_blank,
        show_input_message,
        show_error_message,
        error_title: None,
        error: None,
        prompt_title: None,
        prompt: None,
        formula1: condition,
        formula2: None,
    };
    Some((name, def))
}

fn parse_ods_table(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    sheet_id: u32,
    store: &mut IrStore,
    validations: &HashMap<String, ValidationDef>,
    limits: &dyn OdfLimitCounter,
) -> Result<Worksheet, ParseError> {
    if limits.fast_mode() {
        return parse_ods_table_fast(reader, start, sheet_id, store, validations, limits);
    }
    let name = attr_value(start, b"table:name").unwrap_or_else(|| format!("Sheet{sheet_id}"));
    let mut worksheet = Worksheet::new(name, sheet_id);
    let mut buf = Vec::new();
    let mut row_idx: u32 = 0;
    let mut style_map: HashMap<String, u32> = HashMap::new();
    let mut next_style_id = 1u32;
    let mut validation_ranges: HashMap<String, Vec<String>> = HashMap::new();
    let mut shapes: Vec<NodeId> = Vec::new();
    let mut cell_values: HashMap<(u32, u32), CellValue> = HashMap::new();
    let mut formula_cells: Vec<(NodeId, u32, u32, String)> = Vec::new();
    let mut formula_map: HashMap<(u32, u32), String> = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    let row_repeat = attr_value(&e, b"table:number-rows-repeated")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(1);
                    limits.bump_rows(row_repeat as u64)?;
                    let row_cells =
                        parse_ods_row(reader, &e, store, &mut style_map, &mut next_style_id)?;
                    for _ in 0..row_repeat {
                        let mut col_idx: u32 = 0;
                        for cell_data in &row_cells.cells {
                            for _ in 0..cell_data.col_repeat {
                                if cell_data.is_covered {
                                    col_idx += 1;
                                    continue;
                                }
                                if let Some(range) = cell_data.merge_range(row_idx, col_idx) {
                                    worksheet.merged_cells.push(range);
                                }
                                if cell_data.should_emit() {
                                    limits.bump_cells(1)?;
                                    let reference =
                                        format!("{}{}", column_index_to_name(col_idx), row_idx + 1);
                                    let mut cell = Cell::new(reference.clone(), col_idx, row_idx);
                                    cell.value = cell_data.value.clone();
                                    cell.formula = cell_data.formula.clone();
                                    cell.style_id = cell_data.style_id;
                                    let cell_id = cell.id;
                                    store.insert(IRNode::Cell(cell));
                                    worksheet.cells.push(cell_id);
                                    cell_values.insert((row_idx, col_idx), cell_data.value.clone());
                                    if let Some(formula) = &cell_data.formula {
                                        formula_map
                                            .insert((row_idx, col_idx), formula.text.clone());
                                        formula_cells.push((
                                            cell_id,
                                            row_idx,
                                            col_idx,
                                            formula.text.clone(),
                                        ));
                                    }
                                    if let Some(validation) = &cell_data.validation_name {
                                        validation_ranges
                                            .entry(validation.clone())
                                            .or_default()
                                            .push(reference);
                                    }
                                }
                                col_idx += 1;
                            }
                        }
                        row_idx += 1;
                    }
                }
                b"draw:frame" => {
                    if let Some(shape_id) = parse_draw_frame_spreadsheet(reader, &e, store)? {
                        shapes.push(shape_id);
                    }
                }
                b"table:conditional-formatting" => {
                    if let Some(cf) = parse_ods_conditional_formatting(reader, &e)? {
                        let id = cf.id;
                        store.insert(IRNode::ConditionalFormat(cf));
                        worksheet.conditional_formats.push(id);
                    }
                }
                b"table:filter" | b"table:filter-and" | b"table:filter-or" => {
                    spreadsheet::skip_element(reader, e.name().as_ref())?;
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    let row_repeat = attr_value(&e, b"table:number-rows-repeated")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(1);
                    limits.bump_rows(row_repeat as u64)?;
                    row_idx += row_repeat;
                }
                b"draw:frame" => {
                    if let Some(shape_id) = parse_draw_frame_spreadsheet(reader, &e, store)? {
                        shapes.push(shape_id);
                    }
                }
                b"table:conditional-formatting" => {
                    if let Some(cf) = parse_ods_conditional_formatting_empty(&e)? {
                        let id = cf.id;
                        store.insert(IRNode::ConditionalFormat(cf));
                        worksheet.conditional_formats.push(id);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if !formula_cells.is_empty() {
        evaluate_ods_formulas(
            &worksheet.name,
            &formula_cells,
            store,
            &mut cell_values,
            &formula_map,
        );
    }

    if !validation_ranges.is_empty() {
        for (name, ranges) in validation_ranges {
            let mut validation = if let Some(def) = validations.get(&name) {
                DataValidation {
                    id: NodeId::new(),
                    validation_type: def.validation_type.clone(),
                    operator: def.operator.clone(),
                    allow_blank: def.allow_blank,
                    show_input_message: def.show_input_message,
                    show_error_message: def.show_error_message,
                    error_title: def.error_title.clone(),
                    error: def.error.clone(),
                    prompt_title: def.prompt_title.clone(),
                    prompt: def.prompt.clone(),
                    ranges: ranges.clone(),
                    formula1: def.formula1.clone(),
                    formula2: def.formula2.clone(),
                    span: None,
                }
            } else {
                DataValidation {
                    id: NodeId::new(),
                    validation_type: None,
                    operator: None,
                    allow_blank: false,
                    show_input_message: false,
                    show_error_message: false,
                    error_title: None,
                    error: None,
                    prompt_title: None,
                    prompt: None,
                    ranges: ranges.clone(),
                    formula1: None,
                    formula2: None,
                    span: None,
                }
            };
            let validation_id = validation.id;
            store.insert(IRNode::DataValidation(validation));
            worksheet.data_validations.push(validation_id);
        }
    }

    if !shapes.is_empty() {
        let mut drawing = WorksheetDrawing::new();
        drawing.shapes = shapes;
        let drawing_id = drawing.id;
        store.insert(IRNode::WorksheetDrawing(drawing));
        worksheet.drawings.push(drawing_id);
    }

    Ok(worksheet)
}

fn parse_ods_table_fast(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    sheet_id: u32,
    store: &mut IrStore,
    validations: &HashMap<String, ValidationDef>,
    limits: &dyn OdfLimitCounter,
) -> Result<Worksheet, ParseError> {
    let name = attr_value(start, b"table:name").unwrap_or_else(|| format!("Sheet{sheet_id}"));
    let mut worksheet = Worksheet::new(name, sheet_id);
    let mut buf = Vec::new();
    let mut row_idx: u32 = 0;
    let mut style_map: HashMap<String, u32> = HashMap::new();
    let mut next_style_id = 1u32;
    let mut validation_ranges: HashMap<String, Vec<String>> = HashMap::new();
    let sample_rows = limits.sample_rows();
    let sample_cols = limits.sample_cols();
    let sample_enabled = sample_rows > 0 && sample_cols > 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    let row_repeat = attr_value(&e, b"table:number-rows-repeated")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(1);
                    limits.bump_rows(row_repeat as u64)?;

                    if sample_enabled && row_idx < sample_rows {
                        let row_cells = parse_ods_row_sample(
                            reader,
                            &e,
                            store,
                            &mut style_map,
                            &mut next_style_id,
                            sample_cols,
                        )?;
                        let remaining = sample_rows.saturating_sub(row_idx);
                        let repeat = row_repeat.min(remaining);
                        for _ in 0..repeat {
                            let mut col_idx: u32 = 0;
                            for cell_data in &row_cells.cells {
                                for _ in 0..cell_data.col_repeat {
                                    if col_idx >= sample_cols {
                                        break;
                                    }
                                    if cell_data.is_covered {
                                        col_idx += 1;
                                        continue;
                                    }
                                    if let Some(range) = cell_data.merge_range(row_idx, col_idx) {
                                        worksheet.merged_cells.push(range);
                                    }
                                    if cell_data.should_emit() {
                                        limits.bump_cells(1)?;
                                        let reference = format!(
                                            "{}{}",
                                            column_index_to_name(col_idx),
                                            row_idx + 1
                                        );
                                        let mut cell =
                                            Cell::new(reference.clone(), col_idx, row_idx);
                                        cell.value = cell_data.value.clone();
                                        cell.formula = cell_data.formula.clone();
                                        cell.style_id = cell_data.style_id;
                                        let cell_id = cell.id;
                                        store.insert(IRNode::Cell(cell));
                                        worksheet.cells.push(cell_id);
                                        if let Some(validation) = &cell_data.validation_name {
                                            validation_ranges
                                                .entry(validation.clone())
                                                .or_default()
                                                .push(reference);
                                        }
                                    }
                                    col_idx += 1;
                                }
                                if col_idx >= sample_cols {
                                    break;
                                }
                            }
                            row_idx += 1;
                        }
                        if row_repeat > repeat {
                            row_idx += row_repeat - repeat;
                        }
                    } else {
                        spreadsheet::skip_element(reader, e.name().as_ref())?;
                        row_idx += row_repeat;
                    }
                }
                b"draw:frame" => {
                    spreadsheet::skip_element(reader, e.name().as_ref())?;
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    let row_repeat = attr_value(&e, b"table:number-rows-repeated")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(1);
                    limits.bump_rows(row_repeat as u64)?;
                    row_idx += row_repeat;
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if !validation_ranges.is_empty() {
        for (name, ranges) in validation_ranges {
            let mut validation = if let Some(def) = validations.get(&name) {
                DataValidation {
                    id: NodeId::new(),
                    validation_type: def.validation_type.clone(),
                    operator: def.operator.clone(),
                    allow_blank: def.allow_blank,
                    show_input_message: def.show_input_message,
                    show_error_message: def.show_error_message,
                    error_title: def.error_title.clone(),
                    error: def.error.clone(),
                    prompt_title: def.prompt_title.clone(),
                    prompt: def.prompt.clone(),
                    ranges: ranges.clone(),
                    formula1: def.formula1.clone(),
                    formula2: def.formula2.clone(),
                    span: None,
                }
            } else {
                DataValidation {
                    id: NodeId::new(),
                    validation_type: None,
                    operator: None,
                    allow_blank: false,
                    show_input_message: false,
                    show_error_message: false,
                    error_title: None,
                    error: None,
                    prompt_title: None,
                    prompt: None,
                    ranges: ranges.clone(),
                    formula1: None,
                    formula2: None,
                    span: None,
                }
            };
            validation.span = Some(SourceSpan::new("content.xml"));
            let validation_id = validation.id;
            store.insert(IRNode::DataValidation(validation));
            worksheet.data_validations.push(validation_id);
        }
    }

    Ok(worksheet)
}

fn parse_ods_conditional_formatting(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
) -> Result<Option<ConditionalFormat>, ParseError> {
    let mut cf = ConditionalFormat {
        id: NodeId::new(),
        ranges: Vec::new(),
        rules: Vec::new(),
        span: Some(SourceSpan::new("content.xml")),
    };
    if let Some(ranges) = attr_value(start, b"table:target-range-address")
        .or_else(|| attr_value(start, b"table:cell-range-address"))
    {
        cf.ranges = ranges.split_whitespace().map(|s| s.to_string()).collect();
    }

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"table:conditional-format" {
                    let rule = build_ods_conditional_rule(&e);
                    cf.rules.push(rule);
                    spreadsheet::skip_element(reader, e.name().as_ref())?;
                }
            }
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"table:conditional-format" {
                    let rule = build_ods_conditional_rule(&e);
                    cf.rules.push(rule);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:conditional-formatting" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if cf.rules.is_empty() && cf.ranges.is_empty() {
        Ok(None)
    } else {
        Ok(Some(cf))
    }
}

fn parse_ods_conditional_formatting_empty(
    start: &BytesStart<'_>,
) -> Result<Option<ConditionalFormat>, ParseError> {
    let mut cf = ConditionalFormat {
        id: NodeId::new(),
        ranges: Vec::new(),
        rules: Vec::new(),
        span: Some(SourceSpan::new("content.xml")),
    };
    if let Some(ranges) = attr_value(start, b"table:target-range-address")
        .or_else(|| attr_value(start, b"table:cell-range-address"))
    {
        cf.ranges = ranges.split_whitespace().map(|s| s.to_string()).collect();
    }
    if cf.rules.is_empty() && cf.ranges.is_empty() {
        Ok(None)
    } else {
        Ok(Some(cf))
    }
}

fn build_ods_conditional_rule(start: &BytesStart<'_>) -> ConditionalRule {
    let mut rule = ConditionalRule {
        rule_type: "odf-condition".to_string(),
        priority: None,
        operator: None,
        formulae: Vec::new(),
    };
    rule.priority = attr_value(start, b"table:priority").and_then(|v| v.parse::<u32>().ok());
    if let Some(condition) = attr_value(start, b"table:condition") {
        rule.operator = parse_odf_condition_operator(&condition);
        rule.formulae.push(condition);
    }
    if let Some(style_name) = attr_value(start, b"table:apply-style-name") {
        rule.formulae.push(format!("apply-style:{}", style_name));
    }
    rule
}

fn parse_odf_condition_operator(condition: &str) -> Option<String> {
    let lower = condition.to_ascii_lowercase();
    if let Some(idx) = lower.find("cell-content-is-") {
        let rest = &lower[idx + "cell-content-is-".len()..];
        let op = rest.split('(').next().unwrap_or(rest);
        return Some(op.to_string());
    }
    if let Some(idx) = lower.find("is-true-formula") {
        let _ = idx;
        return Some("true-formula".to_string());
    }
    if let Some(idx) = lower.find("formula-is") {
        let _ = idx;
        return Some("formula".to_string());
    }
    None
}

#[derive(Debug, Clone)]
struct OdsRow {
    cells: Vec<OdsCellData>,
}

#[derive(Debug, Clone)]
struct OdsCellData {
    value: CellValue,
    formula: Option<CellFormula>,
    style_id: Option<u32>,
    col_repeat: u32,
    validation_name: Option<String>,
    col_span: Option<u32>,
    row_span: Option<u32>,
    is_covered: bool,
}

impl OdsCellData {
    fn should_emit(&self) -> bool {
        !matches!(self.value, CellValue::Empty) || self.formula.is_some()
    }

    fn merge_range(&self, row: u32, col: u32) -> Option<MergedCellRange> {
        let col_span = self.col_span.unwrap_or(1);
        let row_span = self.row_span.unwrap_or(1);
        if col_span > 1 || row_span > 1 {
            Some(MergedCellRange {
                start_col: col,
                start_row: row,
                end_col: col + col_span - 1,
                end_row: row + row_span - 1,
            })
        } else {
            None
        }
    }
}

fn parse_ods_row(
    reader: &mut OdfReader<'_>,
    _start: &BytesStart<'_>,
    store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<OdsRow, ParseError> {
    let mut buf = Vec::new();
    let mut cells = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table-cell" => {
                    let cell = parse_ods_cell(reader, &e, store, style_map, next_style_id)?;
                    cells.push(cell);
                }
                b"table:covered-table-cell" => {
                    let cell = parse_ods_covered_cell(reader, &e)?;
                    cells.push(cell);
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-cell" => {
                    let cell = parse_ods_cell_empty(&e, style_map, next_style_id)?;
                    cells.push(cell);
                }
                b"table:covered-table-cell" => {
                    let cell = parse_ods_covered_cell_empty(&e)?;
                    cells.push(cell);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table-row" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(OdsRow { cells })
}

fn parse_ods_row_sample(
    reader: &mut OdfReader<'_>,
    _start: &BytesStart<'_>,
    store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
    sample_cols: u32,
) -> Result<OdsRow, ParseError> {
    let mut buf = Vec::new();
    let mut cells = Vec::new();
    let mut col_idx: u32 = 0;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table-cell" => {
                    if col_idx >= sample_cols {
                        let repeat = attr_value(&e, b"table:number-columns-repeated")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        col_idx = col_idx.saturating_add(repeat);
                        spreadsheet::skip_element(reader, e.name().as_ref())?;
                    } else {
                        let cell = parse_ods_cell(reader, &e, store, style_map, next_style_id)?;
                        col_idx = col_idx.saturating_add(cell.col_repeat);
                        cells.push(cell);
                    }
                }
                b"table:covered-table-cell" => {
                    if col_idx >= sample_cols {
                        let repeat = attr_value(&e, b"table:number-columns-repeated")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        col_idx = col_idx.saturating_add(repeat);
                        spreadsheet::skip_element(reader, e.name().as_ref())?;
                    } else {
                        let cell = parse_ods_covered_cell(reader, &e)?;
                        col_idx = col_idx.saturating_add(cell.col_repeat);
                        cells.push(cell);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-cell" => {
                    if col_idx < sample_cols {
                        let cell = parse_ods_cell_empty(&e, style_map, next_style_id)?;
                        col_idx = col_idx.saturating_add(cell.col_repeat);
                        cells.push(cell);
                    } else {
                        let repeat = attr_value(&e, b"table:number-columns-repeated")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        col_idx = col_idx.saturating_add(repeat);
                    }
                }
                b"table:covered-table-cell" => {
                    if col_idx < sample_cols {
                        let cell = parse_ods_covered_cell_empty(&e)?;
                        col_idx = col_idx.saturating_add(cell.col_repeat);
                        cells.push(cell);
                    } else {
                        let repeat = attr_value(&e, b"table:number-columns-repeated")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        col_idx = col_idx.saturating_add(repeat);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table-row" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(OdsRow { cells })
}

fn parse_ods_cell(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    _store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<OdsCellData, ParseError> {
    let mut value_type = attr_value(start, b"table:cell-value-type")
        .or_else(|| attr_value(start, b"office:value-type"));
    let mut value_attr =
        attr_value(start, b"table:cell-value").or_else(|| attr_value(start, b"office:value"));
    let date_value = attr_value(start, b"table:date-value");
    let time_value = attr_value(start, b"table:time-value");
    let formula_attr = attr_value(start, b"table:formula");
    let col_repeat = attr_value(start, b"table:number-columns-repeated")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let col_span =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok());
    let row_span =
        attr_value(start, b"table:number-rows-spanned").and_then(|v| v.parse::<u32>().ok());
    let validation_name = attr_value(start, b"table:content-validation-name");

    let style_id = attr_value(start, b"table:style-name").and_then(|name| {
        if let Some(id) = style_map.get(&name) {
            Some(*id)
        } else {
            let id = *next_style_id;
            *next_style_id += 1;
            style_map.insert(name, id);
            Some(id)
        }
    });

    let mut text = String::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:p" => {
                    let para = parse_text_element(reader, e.name().as_ref())?;
                    if !text.is_empty() && !para.is_empty() {
                        text.push('\n');
                    }
                    text.push_str(&para);
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                let chunk = e.unescape().unwrap_or_default();
                text.push_str(&chunk);
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table-cell" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if value_type.is_none() && !text.is_empty() {
        value_type = Some("string".to_string());
    }

    if value_attr.is_none() {
        value_attr = date_value.or(time_value).or_else(|| {
            if !text.is_empty() {
                Some(text.clone())
            } else {
                None
            }
        });
    }

    let value = match value_type.as_deref() {
        Some("string") | Some("text") => {
            if !text.is_empty() {
                CellValue::String(text.clone())
            } else {
                CellValue::String(value_attr.unwrap_or_default())
            }
        }
        Some("float") | Some("currency") | Some("percentage") => value_attr
            .as_deref()
            .and_then(|v| v.parse::<f64>().ok())
            .map(CellValue::Number)
            .unwrap_or_else(|| CellValue::String(text.clone())),
        Some("boolean") => {
            let v = value_attr.unwrap_or_else(|| text.clone());
            CellValue::Boolean(v == "true" || v == "1")
        }
        Some("date") | Some("time") => {
            CellValue::String(value_attr.unwrap_or_else(|| text.clone()))
        }
        _ => {
            if !text.is_empty() {
                CellValue::String(text.clone())
            } else {
                CellValue::Empty
            }
        }
    };

    let formula = formula_attr.and_then(|formula| {
        let mut text = formula;
        if let Some(stripped) = text.strip_prefix("of:=") {
            text = stripped.to_string();
        } else if let Some(stripped) = text.strip_prefix("of:") {
            text = stripped.to_string();
        } else if let Some(stripped) = text.strip_prefix('=') {
            text = stripped.to_string();
        }
        Some(CellFormula {
            text,
            formula_type: docir_core::ir::FormulaType::Normal,
            shared_index: None,
            shared_ref: None,
            is_array: false,
            array_ref: None,
        })
    });

    Ok(OdsCellData {
        value,
        formula,
        style_id,
        col_repeat,
        validation_name,
        col_span,
        row_span,
        is_covered: false,
    })
}

fn eval_simple_formula(formula: &str) -> Option<f64> {
    if formula.is_empty() {
        return None;
    }
    if formula
        .chars()
        .any(|c| c.is_alphabetic() || c == '[' || c == ':' || c == ';')
    {
        return None;
    }
    let mut chars = formula.chars().peekable();
    parse_expr(&mut chars)
}

fn parse_expr(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> Option<f64> {
    let mut value = parse_term(chars)?;
    loop {
        skip_ws(chars);
        match chars.peek().copied() {
            Some('+') => {
                chars.next();
                value += parse_term(chars)?;
            }
            Some('-') => {
                chars.next();
                value -= parse_term(chars)?;
            }
            _ => break,
        }
    }
    Some(value)
}

fn parse_term(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> Option<f64> {
    let mut value = parse_factor(chars)?;
    loop {
        skip_ws(chars);
        match chars.peek().copied() {
            Some('*') => {
                chars.next();
                value *= parse_factor(chars)?;
            }
            Some('/') => {
                chars.next();
                let denom = parse_factor(chars)?;
                if denom == 0.0 {
                    return None;
                }
                value /= denom;
            }
            _ => break,
        }
    }
    Some(value)
}

fn parse_factor(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> Option<f64> {
    skip_ws(chars);
    if chars.peek() == Some(&'(') {
        chars.next();
        let value = parse_expr(chars)?;
        skip_ws(chars);
        if chars.peek() == Some(&')') {
            chars.next();
            return Some(value);
        }
        return None;
    }
    parse_number(chars)
}

fn parse_number(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> Option<f64> {
    skip_ws(chars);
    let mut s = String::new();
    while let Some(&c) = chars.peek() {
        if c.is_ascii_digit() || c == '.' {
            s.push(c);
            chars.next();
        } else if c == '-' && s.is_empty() {
            s.push(c);
            chars.next();
        } else {
            break;
        }
    }
    if s.is_empty() {
        None
    } else {
        s.parse::<f64>().ok()
    }
}

fn skip_ws(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) {
    while matches!(chars.peek(), Some(' ' | '\t' | '\n' | '\r')) {
        chars.next();
    }
}

#[derive(Debug, Clone)]
struct CellRef {
    sheet: Option<String>,
    row: u32,
    col: u32,
}

#[derive(Debug, Clone)]
struct CellRange {
    start: CellRef,
    end: CellRef,
}

#[derive(Debug, Clone)]
enum FormulaToken {
    Number(f64),
    Ident(String),
    Ref(CellRef),
    Range(CellRange),
    Plus,
    Minus,
    Star,
    Slash,
    LParen,
    RParen,
    Comma,
    End,
}

struct FormulaEvalContext<'a> {
    sheet_name: &'a str,
    values: HashMap<(u32, u32), CellValue>,
    formulas: &'a HashMap<(u32, u32), String>,
    cache: HashMap<(u32, u32), Option<f64>>,
    stack: Vec<(u32, u32)>,
}

impl<'a> FormulaEvalContext<'a> {
    fn new(
        sheet_name: &'a str,
        values: HashMap<(u32, u32), CellValue>,
        formulas: &'a HashMap<(u32, u32), String>,
    ) -> Self {
        Self {
            sheet_name,
            values,
            formulas,
            cache: HashMap::new(),
            stack: Vec::new(),
        }
    }

    fn eval_formula(&mut self, formula: &str) -> Option<f64> {
        let tokens = tokenize_formula(formula);
        let mut parser = FormulaParser::new(tokens, self);
        parser.parse_expression()
    }

    fn resolve_ref(&mut self, reference: &CellRef) -> Option<f64> {
        if let Some(sheet) = reference.sheet.as_deref() {
            if !sheet.eq_ignore_ascii_case(self.sheet_name) {
                return None;
            }
        }
        let key = (reference.row, reference.col);
        if let Some(value) = self.cache.get(&key) {
            return *value;
        }
        if self.stack.contains(&key) {
            self.cache.insert(key, None);
            return None;
        }
        if let Some(value) = self.values.get(&key) {
            if let Some(number) = cell_value_to_number(value) {
                self.cache.insert(key, Some(number));
                return Some(number);
            }
        }
        let formula_text = self.formulas.get(&key)?.clone();
        self.stack.push(key);
        let result = self.eval_formula(&formula_text);
        self.stack.pop();
        if let Some(number) = result {
            self.values.insert(key, CellValue::Number(number));
        }
        self.cache.insert(key, result);
        result
    }

    fn resolve_range(&mut self, range: &CellRange) -> Option<Vec<f64>> {
        if let Some(sheet) = range.start.sheet.as_deref() {
            if !sheet.eq_ignore_ascii_case(self.sheet_name) {
                return None;
            }
        }
        if let Some(sheet) = range.end.sheet.as_deref() {
            if !sheet.eq_ignore_ascii_case(self.sheet_name) {
                return None;
            }
        }
        let row_start = range.start.row.min(range.end.row);
        let row_end = range.start.row.max(range.end.row);
        let col_start = range.start.col.min(range.end.col);
        let col_end = range.start.col.max(range.end.col);
        let total = (row_end - row_start + 1) as u64 * (col_end - col_start + 1) as u64;
        if total > 1_000_000 {
            return None;
        }
        let mut values = Vec::new();
        for row in row_start..=row_end {
            for col in col_start..=col_end {
                let reference = CellRef {
                    sheet: None,
                    row,
                    col,
                };
                if let Some(number) = self.resolve_ref(&reference) {
                    values.push(number);
                }
            }
        }
        Some(values)
    }
}

struct FormulaParser<'a, 'b> {
    tokens: Vec<FormulaToken>,
    pos: usize,
    ctx: &'a mut FormulaEvalContext<'b>,
}

impl<'a, 'b> FormulaParser<'a, 'b> {
    fn new(tokens: Vec<FormulaToken>, ctx: &'a mut FormulaEvalContext<'b>) -> Self {
        Self {
            tokens,
            pos: 0,
            ctx,
        }
    }

    fn parse_expression(&mut self) -> Option<f64> {
        let mut value = self.parse_term()?;
        loop {
            match self.peek() {
                FormulaToken::Plus => {
                    self.next();
                    value += self.parse_term()?;
                }
                FormulaToken::Minus => {
                    self.next();
                    value -= self.parse_term()?;
                }
                _ => break,
            }
        }
        Some(value)
    }

    fn parse_term(&mut self) -> Option<f64> {
        let mut value = self.parse_factor()?;
        loop {
            match self.peek() {
                FormulaToken::Star => {
                    self.next();
                    value *= self.parse_factor()?;
                }
                FormulaToken::Slash => {
                    self.next();
                    let denom = self.parse_factor()?;
                    if denom == 0.0 {
                        return None;
                    }
                    value /= denom;
                }
                _ => break,
            }
        }
        Some(value)
    }

    fn parse_factor(&mut self) -> Option<f64> {
        let token = self.peek().clone();
        match token {
            FormulaToken::Minus => {
                self.next();
                self.parse_factor().map(|v| -v)
            }
            FormulaToken::Number(value) => {
                self.next();
                Some(value)
            }
            FormulaToken::Ref(reference) => {
                self.next();
                self.ctx.resolve_ref(&reference)
            }
            FormulaToken::Range(range) => {
                self.next();
                let values = self.ctx.resolve_range(&range)?;
                Some(values.iter().sum())
            }
            FormulaToken::Ident(name) => {
                self.next();
                if matches!(self.peek(), FormulaToken::LParen) {
                    self.next();
                    let values = self.parse_function_args()?;
                    if !matches!(self.peek(), FormulaToken::RParen) {
                        return None;
                    }
                    self.next();
                    eval_formula_function(&name, &values)
                } else {
                    None
                }
            }
            FormulaToken::LParen => {
                self.next();
                let value = self.parse_expression()?;
                if !matches!(self.peek(), FormulaToken::RParen) {
                    return None;
                }
                self.next();
                Some(value)
            }
            _ => None,
        }
    }

    fn parse_function_args(&mut self) -> Option<Vec<f64>> {
        let mut values = Vec::new();
        if matches!(self.peek(), FormulaToken::RParen) {
            return Some(values);
        }
        loop {
            if matches!(self.peek(), FormulaToken::Range(_)) {
                if let FormulaToken::Range(range) = self.next().clone() {
                    let range_values = self.ctx.resolve_range(&range)?;
                    values.extend(range_values);
                }
            } else {
                let value = self.parse_expression()?;
                values.push(value);
            }
            match self.peek() {
                FormulaToken::Comma => {
                    self.next();
                }
                FormulaToken::RParen => break,
                _ => return None,
            }
        }
        Some(values)
    }

    fn peek(&self) -> &FormulaToken {
        self.tokens.get(self.pos).unwrap_or(&FormulaToken::End)
    }

    fn next(&mut self) -> &FormulaToken {
        let token = self.tokens.get(self.pos).unwrap_or(&FormulaToken::End);
        self.pos += 1;
        token
    }
}

fn tokenize_formula(formula: &str) -> Vec<FormulaToken> {
    let mut tokens = Vec::new();
    let mut chars = formula.trim().chars().peekable();
    while let Some(&ch) = chars.peek() {
        match ch {
            ' ' | '\t' | '\n' | '\r' => {
                chars.next();
            }
            '+' => {
                chars.next();
                tokens.push(FormulaToken::Plus);
            }
            '-' => {
                chars.next();
                tokens.push(FormulaToken::Minus);
            }
            '*' => {
                chars.next();
                tokens.push(FormulaToken::Star);
            }
            '/' => {
                chars.next();
                tokens.push(FormulaToken::Slash);
            }
            '(' => {
                chars.next();
                tokens.push(FormulaToken::LParen);
            }
            ')' => {
                chars.next();
                tokens.push(FormulaToken::RParen);
            }
            ',' | ';' => {
                chars.next();
                tokens.push(FormulaToken::Comma);
            }
            '[' => {
                chars.next();
                let mut buffer = String::new();
                while let Some(c) = chars.next() {
                    if c == ']' {
                        break;
                    }
                    buffer.push(c);
                }
                if let Some(token) = parse_bracket_reference(&buffer) {
                    tokens.push(token);
                }
            }
            _ => {
                if ch.is_ascii_digit() || ch == '.' {
                    let mut num = String::new();
                    while let Some(&c) = chars.peek() {
                        if c.is_ascii_digit() || c == '.' {
                            num.push(c);
                            chars.next();
                        } else {
                            break;
                        }
                    }
                    if let Ok(value) = num.parse::<f64>() {
                        tokens.push(FormulaToken::Number(value));
                    }
                } else if ch.is_ascii_alphabetic() || ch == '_' {
                    let mut ident = String::new();
                    while let Some(&c) = chars.peek() {
                        if c.is_ascii_alphanumeric() || c == '_' || c == '.' || c == '$' {
                            ident.push(c);
                            chars.next();
                        } else {
                            break;
                        }
                    }
                    if let Some(reference) = parse_simple_reference(&ident) {
                        tokens.push(FormulaToken::Ref(reference));
                    } else {
                        tokens.push(FormulaToken::Ident(ident));
                    }
                } else {
                    chars.next();
                }
            }
        }
    }
    tokens.push(FormulaToken::End);
    tokens
}

fn parse_bracket_reference(input: &str) -> Option<FormulaToken> {
    let trimmed = input.trim();
    if let Some((start, end)) = trimmed.split_once(':') {
        let start_ref = parse_sheeted_cell(start)?;
        let end_ref = parse_sheeted_cell(end)?;
        return Some(FormulaToken::Range(CellRange {
            start: start_ref,
            end: end_ref,
        }));
    }
    parse_sheeted_cell(trimmed).map(FormulaToken::Ref)
}

fn parse_simple_reference(input: &str) -> Option<CellRef> {
    if input.chars().any(|c| c.is_ascii_digit()) && input.chars().any(|c| c.is_ascii_alphabetic()) {
        parse_sheeted_cell(input)
    } else {
        None
    }
}

fn parse_sheeted_cell(input: &str) -> Option<CellRef> {
    let trimmed = input.trim().trim_start_matches('.');
    let mut sheet: Option<String> = None;
    let mut cell_part = trimmed;
    if let Some((sheet_part, cell)) = trimmed.rsplit_once('.') {
        if !sheet_part.is_empty() && !cell.is_empty() {
            let sheet_name = sheet_part.trim_matches('\'').replace('$', "");
            if !sheet_name.is_empty() {
                sheet = Some(sheet_name);
            }
            cell_part = cell;
        }
    }
    let cell = parse_cell_ref(cell_part)?;
    Some(CellRef {
        sheet,
        row: cell.row,
        col: cell.col,
    })
}

fn parse_cell_ref(input: &str) -> Option<CellRef> {
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in input.chars() {
        if ch.is_ascii_alphabetic() {
            letters.push(ch);
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else if ch == '$' {
            continue;
        } else {
            break;
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    let col = column_name_to_index(&letters)?;
    let row = digits.parse::<u32>().ok()?.saturating_sub(1);
    Some(CellRef {
        sheet: None,
        row,
        col,
    })
}

fn column_name_to_index(name: &str) -> Option<u32> {
    let mut index: u32 = 0;
    for ch in name.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        index = index * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    Some(index.saturating_sub(1))
}

fn cell_value_to_number(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Number(num) => Some(*num),
        CellValue::Boolean(v) => Some(if *v { 1.0 } else { 0.0 }),
        CellValue::String(s) => s.parse::<f64>().ok(),
        _ => None,
    }
}

fn eval_formula_function(name: &str, values: &[f64]) -> Option<f64> {
    let upper = name.to_ascii_uppercase();
    match upper.as_str() {
        "SUM" => Some(values.iter().sum()),
        "AVERAGE" => {
            if values.is_empty() {
                None
            } else {
                Some(values.iter().sum::<f64>() / values.len() as f64)
            }
        }
        "MIN" => values.iter().copied().reduce(f64::min),
        "MAX" => values.iter().copied().reduce(f64::max),
        "COUNT" => Some(values.len() as f64),
        _ => None,
    }
}

fn evaluate_ods_formulas(
    sheet_name: &str,
    formula_cells: &[(NodeId, u32, u32, String)],
    store: &mut IrStore,
    cell_values: &mut HashMap<(u32, u32), CellValue>,
    formula_map: &HashMap<(u32, u32), String>,
) {
    let mut ctx = FormulaEvalContext::new(sheet_name, cell_values.clone(), formula_map);
    for (cell_id, row, col, formula) in formula_cells {
        if let Some(IRNode::Cell(cell)) = store.get_mut(*cell_id) {
            if matches!(cell.value, CellValue::Empty) {
                if let Some(value) = ctx.eval_formula(formula) {
                    cell.value = CellValue::Number(value);
                    cell_values.insert((*row, *col), CellValue::Number(value));
                }
            }
        }
    }
}

fn parse_ods_cell_empty(
    start: &BytesStart<'_>,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<OdsCellData, ParseError> {
    let value_type = attr_value(start, b"table:cell-value-type")
        .or_else(|| attr_value(start, b"office:value-type"));
    let value_attr = attr_value(start, b"table:cell-value")
        .or_else(|| attr_value(start, b"office:value"))
        .or_else(|| attr_value(start, b"table:date-value"))
        .or_else(|| attr_value(start, b"table:time-value"));
    let formula_attr = attr_value(start, b"table:formula");
    let col_repeat = attr_value(start, b"table:number-columns-repeated")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let col_span =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok());
    let row_span =
        attr_value(start, b"table:number-rows-spanned").and_then(|v| v.parse::<u32>().ok());
    let validation_name = attr_value(start, b"table:content-validation-name");
    let style_id = attr_value(start, b"table:style-name").and_then(|name| {
        if let Some(id) = style_map.get(&name) {
            Some(*id)
        } else {
            let id = *next_style_id;
            *next_style_id += 1;
            style_map.insert(name, id);
            Some(id)
        }
    });

    let value = match value_type.as_deref() {
        Some("float") | Some("currency") | Some("percentage") => value_attr
            .as_deref()
            .and_then(|v| v.parse::<f64>().ok())
            .map(CellValue::Number)
            .unwrap_or(CellValue::Empty),
        Some("boolean") => {
            let v = value_attr.unwrap_or_default();
            if v.is_empty() {
                CellValue::Empty
            } else {
                CellValue::Boolean(v == "true" || v == "1")
            }
        }
        Some("string") | Some("text") => value_attr
            .map(CellValue::String)
            .unwrap_or(CellValue::Empty),
        Some("date") | Some("time") => value_attr
            .map(CellValue::String)
            .unwrap_or(CellValue::Empty),
        _ => CellValue::Empty,
    };

    let formula = formula_attr.and_then(|formula| {
        let mut text = formula;
        if let Some(stripped) = text.strip_prefix("of:=") {
            text = stripped.to_string();
        } else if let Some(stripped) = text.strip_prefix("of:") {
            text = stripped.to_string();
        } else if let Some(stripped) = text.strip_prefix('=') {
            text = stripped.to_string();
        }
        Some(CellFormula {
            text,
            formula_type: docir_core::ir::FormulaType::Normal,
            shared_index: None,
            shared_ref: None,
            is_array: false,
            array_ref: None,
        })
    });

    Ok(OdsCellData {
        value,
        formula,
        style_id,
        col_repeat,
        validation_name,
        col_span,
        row_span,
        is_covered: false,
    })
}

fn parse_ods_covered_cell(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
) -> Result<OdsCellData, ParseError> {
    let col_repeat = attr_value(start, b"table:number-columns-repeated")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let col_span =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok());
    let row_span =
        attr_value(start, b"table:number-rows-spanned").and_then(|v| v.parse::<u32>().ok());
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::End(e)) if e.name().as_ref() == b"table:covered-table-cell" => break,
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(OdsCellData {
        value: CellValue::Empty,
        formula: None,
        style_id: None,
        col_repeat,
        validation_name: None,
        col_span,
        row_span,
        is_covered: true,
    })
}

fn parse_ods_covered_cell_empty(start: &BytesStart<'_>) -> Result<OdsCellData, ParseError> {
    let col_repeat = attr_value(start, b"table:number-columns-repeated")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let col_span =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok());
    let row_span =
        attr_value(start, b"table:number-rows-spanned").and_then(|v| v.parse::<u32>().ok());
    Ok(OdsCellData {
        value: CellValue::Empty,
        formula: None,
        style_id: None,
        col_repeat,
        validation_name: None,
        col_span,
        row_span,
        is_covered: true,
    })
}

fn parse_text_element(reader: &mut OdfReader<'_>, end_name: &[u8]) -> Result<String, ParseError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:s" => {
                    let count = attr_value(&e, b"text:c")
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1);
                    text.extend(std::iter::repeat(' ').take(count));
                }
                b"text:tab" => text.push('\t'),
                b"text:line-break" => text.push('\n'),
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"text:s" => {
                    let count = attr_value(&e, b"text:c")
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1);
                    text.extend(std::iter::repeat(' ').take(count));
                }
                b"text:tab" => text.push('\t'),
                b"text:line-break" => text.push('\n'),
                _ => {}
            },
            Ok(Event::Text(e)) => {
                let chunk = e.unescape().unwrap_or_default();
                text.push_str(&chunk);
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_name {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

fn parse_draw_frame_spreadsheet(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let transform = parse_frame_transform(start);
    let mut shape_type = ShapeType::Picture;
    let mut media_target: Option<String> = None;
    let mut buf = Vec::new();
    let mut has_shape = false;
    let mut name = attr_value(start, b"draw:name");
    let mut chart_id: Option<NodeId> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"draw:image" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href);
                        shape_type = ShapeType::Picture;
                        has_shape = true;
                    }
                }
                b"chart:chart" => {
                    shape_type = ShapeType::Chart;
                    has_shape = true;
                    let chart = parse_odf_chart(reader, &e)?;
                    let id = chart.id;
                    store.insert(IRNode::ChartData(chart));
                    chart_id = Some(id);
                }
                b"draw:plugin" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                        shape_type = classify_media_shape(&href);
                        has_shape = true;
                    }
                }
                b"draw:object" | b"draw:object-ole" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                    }
                    shape_type = ShapeType::OleObject;
                    has_shape = true;
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"draw:image" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href);
                        shape_type = ShapeType::Picture;
                        has_shape = true;
                    }
                }
                b"chart:chart" => {
                    shape_type = ShapeType::Chart;
                    has_shape = true;
                    let mut chart = ChartData::new();
                    chart.chart_type = attr_value(&e, b"chart:class");
                    chart.span = Some(SourceSpan::new("content.xml"));
                    let id = chart.id;
                    store.insert(IRNode::ChartData(chart));
                    chart_id = Some(id);
                }
                b"draw:plugin" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                        shape_type = classify_media_shape(&href);
                        has_shape = true;
                    }
                }
                b"draw:object" | b"draw:object-ole" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        media_target = Some(href.clone());
                    }
                    shape_type = ShapeType::OleObject;
                    has_shape = true;
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"draw:frame" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if has_shape {
        let mut shape = Shape::new(shape_type);
        shape.name = name.take();
        shape.media_target = media_target;
        shape.chart_id = chart_id;
        shape.transform = transform;
        let shape_id = shape.id;
        store.insert(IRNode::Shape(shape));
        Ok(Some(shape_id))
    } else {
        Ok(None)
    }
}

fn column_index_to_name(mut index: u32) -> String {
    let mut name = String::new();
    index += 1;
    while index > 0 {
        let rem = ((index - 1) % 26) as u8;
        name.push((b'A' + rem) as char);
        index = (index - 1) / 26;
    }
    name.chars().rev().collect()
}

#[derive(Debug, Clone, Copy)]
struct ListContext {
    num_id: u32,
    level: u32,
}

fn parse_paragraph(
    reader: &mut OdfReader<'_>,
    end_name: &[u8],
    numbering: Option<NumberingInfo>,
    outline_level: Option<u8>,
    store: &mut IrStore,
    inline_nodes: &mut Vec<NodeId>,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    limits.bump_paragraphs(1)?;
    let mut buf = Vec::new();
    let mut text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:s" => {
                    let count = attr_value(&e, b"text:c")
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1);
                    text.extend(std::iter::repeat(' ').take(count));
                }
                b"text:tab" => text.push('\t'),
                b"text:line-break" => text.push('\n'),
                b"text:bookmark-start" => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let mut bookmark = BookmarkStart::new(name.clone());
                        bookmark.name = Some(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkStart(bookmark));
                        inline_nodes.push(bookmark_id);
                    }
                }
                b"text:bookmark-end" => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let bookmark = BookmarkEnd::new(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkEnd(bookmark));
                        inline_nodes.push(bookmark_id);
                    }
                }
                b"text:date" => {
                    let mut field = Field::new(Some("DATE".to_string()));
                    field.instruction_parsed = Some(FieldInstruction {
                        kind: FieldKind::Date,
                        args: Vec::new(),
                        switches: Vec::new(),
                    });
                    let field_id = field.id;
                    store.insert(IRNode::Field(field));
                    inline_nodes.push(field_id);
                }
                b"text:time" => {
                    let field = Field::new(Some("TIME".to_string()));
                    let field_id = field.id;
                    store.insert(IRNode::Field(field));
                    inline_nodes.push(field_id);
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"text:s" => {
                    let count = attr_value(&e, b"text:c")
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1);
                    text.extend(std::iter::repeat(' ').take(count));
                }
                b"text:tab" => text.push('\t'),
                b"text:line-break" => text.push('\n'),
                b"text:bookmark-start" => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let mut bookmark = BookmarkStart::new(name.clone());
                        bookmark.name = Some(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkStart(bookmark));
                        inline_nodes.push(bookmark_id);
                    }
                }
                b"text:bookmark-end" => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let bookmark = BookmarkEnd::new(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkEnd(bookmark));
                        inline_nodes.push(bookmark_id);
                    }
                }
                b"text:date" => {
                    let mut field = Field::new(Some("DATE".to_string()));
                    field.instruction_parsed = Some(FieldInstruction {
                        kind: FieldKind::Date,
                        args: Vec::new(),
                        switches: Vec::new(),
                    });
                    let field_id = field.id;
                    store.insert(IRNode::Field(field));
                    inline_nodes.push(field_id);
                }
                b"text:time" => {
                    let field = Field::new(Some("TIME".to_string()));
                    let field_id = field.id;
                    store.insert(IRNode::Field(field));
                    inline_nodes.push(field_id);
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                let chunk = e.unescape().unwrap_or_default();
                text.push_str(&chunk);
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_name {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(text::build_paragraph(
        store,
        &text,
        numbering,
        outline_level,
    ))
}

fn parse_table(
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut buf = Vec::new();
    let mut table = Table::new();
    let mut current_row: Option<TableRow> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    current_row = Some(TableRow::new());
                }
                b"table:table-cell" => {
                    let cell_id = parse_table_cell(reader, &e, store, limits)?;
                    if let Some(row) = current_row.as_mut() {
                        row.cells.push(cell_id);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    let row = TableRow::new();
                    let row_id = row.id;
                    store.insert(IRNode::TableRow(row));
                    table.rows.push(row_id);
                }
                b"table:table-cell" => {
                    let mut cell = TableCell::new();
                    if let Some(span) = attr_value(&e, b"table:number-columns-spanned")
                        .and_then(|v| v.parse::<u32>().ok())
                    {
                        let mut props = TableCellProperties::default();
                        props.grid_span = Some(span);
                        cell.properties = props;
                    }
                    let cell_id = cell.id;
                    store.insert(IRNode::TableCell(cell));
                    if let Some(row) = current_row.as_mut() {
                        row.cells.push(cell_id);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    if let Some(row) = current_row.take() {
                        let row_id = row.id;
                        store.insert(IRNode::TableRow(row));
                        table.rows.push(row_id);
                    }
                }
                b"table:table" => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    let table_id = table.id;
    store.insert(IRNode::Table(table));
    Ok(table_id)
}

fn parse_table_cell(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut cell = TableCell::new();
    if let Some(span) =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok())
    {
        let mut props = TableCellProperties::default();
        props.grid_span = Some(span);
        cell.properties = props;
    }

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:p" => {
                    let paragraph_id = parse_paragraph(
                        reader,
                        e.name().as_ref(),
                        None,
                        None,
                        store,
                        &mut Vec::new(),
                        limits,
                    )?;
                    cell.content.push(paragraph_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table-cell" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    let cell_id = cell.id;
    store.insert(IRNode::TableCell(cell));
    Ok(cell_id)
}

fn parse_annotation(
    reader: &mut OdfReader<'_>,
    comment_id: &str,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut comment = Comment::new(comment_id);
    let mut buf = Vec::new();
    let mut current = None;

    #[derive(Clone, Copy)]
    enum AnnotationField {
        Creator,
        Date,
    }

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"dc:creator" => current = Some(AnnotationField::Creator),
                b"dc:date" => current = Some(AnnotationField::Date),
                b"text:p" => {
                    let paragraph_id = parse_paragraph(
                        reader,
                        e.name().as_ref(),
                        None,
                        None,
                        store,
                        &mut Vec::new(),
                        limits,
                    )?;
                    comment.content.push(paragraph_id);
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if let Some(field) = current {
                    let value = e.unescape().unwrap_or_default().to_string();
                    match field {
                        AnnotationField::Creator => comment.author = Some(value),
                        AnnotationField::Date => comment.date = Some(value),
                    }
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"office:annotation" => break,
                b"dc:creator" | b"dc:date" => current = None,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    let comment_id = comment.id;
    store.insert(IRNode::Comment(comment));
    Ok(comment_id)
}

fn parse_note(
    reader: &mut OdfReader<'_>,
    note_id: &str,
    note_class: &str,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut buf = Vec::new();
    let mut content: Vec<NodeId> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:p" => {
                    let paragraph_id = parse_paragraph(
                        reader,
                        e.name().as_ref(),
                        None,
                        None,
                        store,
                        &mut Vec::new(),
                        limits,
                    )?;
                    content.push(paragraph_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"text:note" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if note_class == "endnote" {
        let mut endnote = Endnote::new(note_id);
        endnote.content = content;
        let id = endnote.id;
        store.insert(IRNode::Endnote(endnote));
        Ok(id)
    } else {
        let mut footnote = Footnote::new(note_id);
        footnote.content = content;
        let id = footnote.id;
        store.insert(IRNode::Footnote(footnote));
        Ok(id)
    }
}

fn parse_draw_frame(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let mut shape = Shape::new(ShapeType::Picture);
    shape.transform = parse_frame_transform(start);
    shape.name = attr_value(start, b"draw:name");
    let mut buf = Vec::new();
    let mut has_shape = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"draw:image" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        shape.media_target = Some(href);
                        shape.shape_type = ShapeType::Picture;
                        has_shape = true;
                    }
                }
                b"draw:object" | b"draw:object-ole" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        shape.media_target = Some(href);
                    }
                    shape.shape_type = ShapeType::OleObject;
                    has_shape = true;
                }
                b"draw:plugin" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        shape.media_target = Some(href.clone());
                        shape.shape_type = classify_media_shape(&href);
                        has_shape = true;
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"draw:frame" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    if has_shape {
        let shape_id = shape.id;
        store.insert(IRNode::Shape(shape));
        Ok(Some(shape_id))
    } else {
        Ok(None)
    }
}

fn parse_tracked_changes(
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<Vec<NodeId>, ParseError> {
    let mut buf = Vec::new();
    let mut revisions = Vec::new();
    let mut current_revision: Option<Revision> = None;
    let mut current_field: Option<ChangeInfoField> = None;

    enum ChangeInfoField {
        Author,
        Date,
    }

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:changed-region" => {
                    current_revision = None;
                }
                b"text:insertion" => {
                    current_revision = Some(Revision::new(RevisionType::Insert));
                }
                b"text:deletion" => {
                    current_revision = Some(Revision::new(RevisionType::Delete));
                }
                b"dc:creator" => current_field = Some(ChangeInfoField::Author),
                b"dc:date" => current_field = Some(ChangeInfoField::Date),
                b"text:p" => {
                    if let Some(rev) = current_revision.as_mut() {
                        let paragraph_id = parse_paragraph(
                            reader,
                            e.name().as_ref(),
                            None,
                            None,
                            store,
                            &mut Vec::new(),
                            limits,
                        )?;
                        rev.content.push(paragraph_id);
                    }
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if let Some(rev) = current_revision.as_mut() {
                    if let Some(field) = &current_field {
                        let value = e.unescape().unwrap_or_default().to_string();
                        match field {
                            ChangeInfoField::Author => rev.author = Some(value),
                            ChangeInfoField::Date => rev.date = Some(value),
                        }
                    }
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"text:insertion" | b"text:deletion" => {
                    if let Some(rev) = current_revision.take() {
                        let id = rev.id;
                        store.insert(IRNode::Revision(rev));
                        revisions.push(id);
                    }
                }
                b"text:tracked-changes" => break,
                b"dc:creator" | b"dc:date" => current_field = None,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(revisions)
}

fn parse_styles(xml: &str) -> Option<StyleSet> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut styles = StyleSet::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"style:style" => {
                    if let Some(mut style) = build_style_from_start(&e, false) {
                        parse_style_properties(&mut reader, &mut style, b"style:style");
                        styles.styles.push(style);
                    }
                }
                b"style:default-style" => {
                    if let Some(mut style) = build_style_from_start(&e, true) {
                        parse_style_properties(&mut reader, &mut style, b"style:default-style");
                        styles.styles.push(style);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"style:style" => {
                    if let Some(style) = build_style_from_start(&e, false) {
                        styles.styles.push(style);
                    }
                }
                b"style:default-style" => {
                    if let Some(style) = build_style_from_start(&e, true) {
                        styles.styles.push(style);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }

    if styles.styles.is_empty() {
        None
    } else {
        Some(styles)
    }
}

fn map_style_family(e: &BytesStart<'_>) -> StyleType {
    match attr_value(e, b"style:family").as_deref() {
        Some("paragraph") => StyleType::Paragraph,
        Some("text") => StyleType::Character,
        Some("table") => StyleType::Table,
        Some("list") => StyleType::Numbering,
        _ => StyleType::Other,
    }
}

fn build_style_from_start(start: &BytesStart<'_>, is_default: bool) -> Option<Style> {
    let style_id = attr_value(start, b"style:name")
        .or_else(|| attr_value(start, b"style:family").map(|f| format!("default:{f}")));
    let style_id = style_id?;
    let mut style = Style {
        style_id,
        name: attr_value(start, b"style:display-name"),
        style_type: map_style_family(start),
        based_on: attr_value(start, b"style:parent-style-name"),
        next: attr_value(start, b"style:next-style-name"),
        is_default,
        run_props: None,
        paragraph_props: None,
        table_props: None,
    };
    if let Some(family) = attr_value(start, b"style:family") {
        if family == "paragraph" || family == "text" {
            style.is_default = is_default
                || attr_value(start, b"style:default")
                    .map(|v| v == "true")
                    .unwrap_or(false);
        }
    }
    Some(style)
}

fn parse_style_properties(reader: &mut Reader<&[u8]>, style: &mut Style, end_name: &[u8]) {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"style:text-properties" => {
                    let mut props = style.run_props.take().unwrap_or_default();
                    if let Some(font) = attr_value(&e, b"fo:font-family")
                        .or_else(|| attr_value(&e, b"style:font-name"))
                    {
                        props.font_family = Some(font);
                    }
                    if let Some(size) =
                        attr_value(&e, b"fo:font-size").and_then(|v| parse_font_size(&v))
                    {
                        props.font_size = Some(size);
                    }
                    if let Some(weight) = attr_value(&e, b"fo:font-weight") {
                        props.bold = Some(weight.eq_ignore_ascii_case("bold"));
                    }
                    if let Some(style_attr) = attr_value(&e, b"fo:font-style") {
                        props.italic = Some(style_attr.eq_ignore_ascii_case("italic"));
                    }
                    if let Some(color) = attr_value(&e, b"fo:color") {
                        props.color = Some(color);
                    }
                    style.run_props = Some(props);
                }
                b"style:paragraph-properties" => {
                    let mut props = style.paragraph_props.take().unwrap_or_default();
                    if let Some(align) =
                        attr_value(&e, b"fo:text-align").and_then(|v| parse_text_alignment(&v))
                    {
                        props.alignment = Some(align);
                    }
                    style.paragraph_props = Some(props);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_name {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

fn parse_font_size(value: &str) -> Option<u32> {
    let trimmed = value.trim();
    let num = trimmed
        .trim_end_matches("pt")
        .trim_end_matches("px")
        .trim_end_matches("cm")
        .trim_end_matches("mm");
    num.parse::<f32>().ok().map(|v| v.round() as u32)
}

fn merge_styles(existing: &mut StyleSet, incoming: &mut StyleSet) {
    let mut seen = existing
        .styles
        .iter()
        .map(|s| s.style_id.clone())
        .collect::<std::collections::HashSet<String>>();
    for style in incoming.styles.drain(..) {
        if seen.insert(style.style_id.clone()) {
            existing.styles.push(style);
        }
    }
}

fn parse_master_pages(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"style:master-page" {
                    if let Some(name) = attr_value(&e, b"style:name") {
                        out.push(name);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn parse_page_layouts(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"style:page-layout" {
                    if let Some(name) = attr_value(&e, b"style:name") {
                        out.push(name);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn parse_odf_headers_footers(
    xml: &str,
    store: &mut IrStore,
    config: &ParserConfig,
) -> Result<(Vec<NodeId>, Vec<NodeId>), ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml.as_bytes()));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut headers = Vec::new();
    let mut footers = Vec::new();

    let limits = OdfLimits::new(config, false);

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"style:header" | b"style:header-left" => {
                    let content = parse_odf_header_footer_block(
                        &mut reader,
                        e.name().as_ref(),
                        store,
                        &limits,
                    )?;
                    let mut header = Header::new();
                    header.content = content;
                    header.span = Some(SourceSpan::new("styles.xml"));
                    let id = header.id;
                    store.insert(IRNode::Header(header));
                    headers.push(id);
                }
                b"style:footer" | b"style:footer-left" => {
                    let content = parse_odf_header_footer_block(
                        &mut reader,
                        e.name().as_ref(),
                        store,
                        &limits,
                    )?;
                    let mut footer = Footer::new();
                    footer.content = content;
                    footer.span = Some(SourceSpan::new("styles.xml"));
                    let id = footer.id;
                    store.insert(IRNode::Footer(footer));
                    footers.push(id);
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "styles.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok((headers, footers))
}

fn parse_odf_header_footer_block(
    reader: &mut OdfReader<'_>,
    end_name: &[u8],
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<Vec<NodeId>, ParseError> {
    let mut buf = Vec::new();
    let mut content = Vec::new();
    let mut list_stack: Vec<ListContext> = Vec::new();
    let mut list_id_map: HashMap<String, u32> = HashMap::new();
    let mut next_list_id = 1u32;
    let mut pending_inline_nodes: Vec<NodeId> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:list" => {
                    let style_name = attr_value(&e, b"text:style-name").unwrap_or_default();
                    let num_id = list_id_map.entry(style_name).or_insert_with(|| {
                        let id = next_list_id;
                        next_list_id += 1;
                        id
                    });
                    let level = list_stack.len() as u32;
                    list_stack.push(ListContext {
                        num_id: *num_id,
                        level,
                    });
                }
                b"text:p" | b"text:h" => {
                    let outline_level =
                        attr_value(&e, b"text:outline-level").and_then(|v| v.parse::<u8>().ok());
                    let numbering = list_stack.last().map(|ctx| NumberingInfo {
                        num_id: ctx.num_id,
                        level: ctx.level,
                        format: None,
                    });
                    let paragraph_id = parse_paragraph(
                        reader,
                        e.name().as_ref(),
                        numbering,
                        outline_level,
                        store,
                        &mut pending_inline_nodes,
                        limits,
                    )?;
                    content.extend(pending_inline_nodes.drain(..));
                    content.push(paragraph_id);
                }
                b"table:table" => {
                    let table_id = parse_table(reader, store, limits)?;
                    content.extend(pending_inline_nodes.drain(..));
                    content.push(table_id);
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"text:list" => {}
                b"text:p" | b"text:h" => {
                    let outline_level =
                        attr_value(&e, b"text:outline-level").and_then(|v| v.parse::<u8>().ok());
                    let numbering = list_stack.last().map(|ctx| NumberingInfo {
                        num_id: ctx.num_id,
                        level: ctx.level,
                        format: None,
                    });
                    let paragraph_id = text::build_paragraph(store, "", numbering, outline_level);
                    content.extend(pending_inline_nodes.drain(..));
                    content.push(paragraph_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"text:list" => {
                    list_stack.pop();
                }
                _ if e.name().as_ref() == end_name => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "styles.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(content)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::DocumentParser;
    use docir_core::security::ThreatIndicatorType;
    use std::io::{Cursor, Write};
    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_odf_zip(mimetype: &str, content_xml: &str, styles_xml: Option<&str>) -> Vec<u8> {
        build_odf_zip_custom(mimetype, content_xml, styles_xml, None, Vec::new())
    }

    fn build_odf_zip_custom(
        mimetype: &str,
        content_xml: &str,
        styles_xml: Option<&str>,
        manifest_xml: Option<&str>,
        extra_files: Vec<(&str, &[u8])>,
    ) -> Vec<u8> {
        let mut buffer = Vec::new();
        let cursor = Cursor::new(&mut buffer);
        let mut zip = ZipWriter::new(cursor);
        let stored =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
        zip.start_file("mimetype", stored).unwrap();
        zip.write_all(mimetype.as_bytes()).unwrap();

        zip.start_file("META-INF/manifest.xml", FileOptions::<()>::default())
            .unwrap();
        if let Some(xml) = manifest_xml {
            zip.write_all(xml.as_bytes()).unwrap();
        } else {
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#,
            )
            .unwrap();
        }

        zip.start_file("content.xml", FileOptions::<()>::default())
            .unwrap();
        zip.write_all(content_xml.as_bytes()).unwrap();

        zip.start_file("meta.xml", FileOptions::<()>::default())
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
  <office:meta>
    <dc:title>Test Doc</dc:title>
    <dc:creator>docir</dc:creator>
  </office:meta>
</office:document-meta>
"#,
        )
        .unwrap();

        if let Some(styles) = styles_xml {
            zip.start_file("styles.xml", FileOptions::<()>::default())
                .unwrap();
            zip.write_all(styles.as_bytes()).unwrap();
        }

        for (path, bytes) in extra_files {
            zip.start_file(path, FileOptions::<()>::default()).unwrap();
            zip.write_all(bytes).unwrap();
        }

        zip.finish().unwrap();
        buffer
    }

    #[test]
    fn test_parse_odt_minimal() {
        let mimetype = "application/vnd.oasis.opendocument.text";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:p>Hello ODF</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
        let zip_data = build_odf_zip(mimetype, content_xml, None);
        let parser = DocumentParser::new();
        let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
        docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
        assert_eq!(parsed.format, DocumentFormat::OdfText);
        let doc = parsed.document().unwrap();
        assert!(!doc.content.is_empty());
    }

    #[test]
    fn test_parse_ods_minimal() {
        let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1" />
      <table:table table:name="Sheet2" />
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
        let zip_data = build_odf_zip(mimetype, content_xml, None);
        let parser = DocumentParser::new();
        let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
        assert_eq!(parsed.format, DocumentFormat::OdfSpreadsheet);
        let doc = parsed.document().unwrap();
        assert_eq!(doc.content.len(), 2);
    }

    #[test]
    fn test_parse_ods_cells_and_validations() {
        let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:spreadsheet>
      <table:content-validations>
        <table:content-validation table:name="val1" table:condition="cell-content-is-between(1,10)" table:allow-empty-cell="true" />
      </table:content-validations>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell table:cell-value-type="float" table:cell-value="3.14" />
          <table:table-cell table:cell-value-type="string">
            <text:p>Hello</text:p>
          </table:table-cell>
          <table:table-cell table:formula="of:=SUM([.A1];[.B1])" table:cell-value-type="float" table:cell-value="6.28" table:content-validation-name="val1" />
        </table:table-row>
        <table:table-row table:number-rows-repeated="2">
          <table:table-cell table:cell-value-type="boolean" table:cell-value="true" table:number-columns-repeated="2" />
        </table:table-row>
        <draw:frame draw:name="Chart1">
          <chart:chart />
        </draw:frame>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
        let zip_data = build_odf_zip(mimetype, content_xml, None);
        let parser = DocumentParser::new();
        let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
        docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
        let doc = parsed.document().unwrap();

        let mut cell_count = 0;
        let mut cell_refs = Vec::new();
        let mut validation_count = 0;
        let mut drawing_count = 0;
        for node in parsed.store.values() {
            match node {
                IRNode::Cell(cell) => {
                    cell_count += 1;
                    cell_refs.push(cell.reference.clone());
                }
                IRNode::DataValidation(_) => validation_count += 1,
                IRNode::WorksheetDrawing(_) => drawing_count += 1,
                _ => {}
            }
        }
        assert!(cell_count >= 5, "cells: {:?}", cell_refs);
        assert_eq!(validation_count, 1);
        assert_eq!(drawing_count, 1);
        assert_eq!(doc.content.len(), 1);
    }

    #[test]
    fn test_parse_odp_minimal() {
        let mimetype = "application/vnd.oasis.opendocument.presentation";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1" />
      <draw:page draw:name="Slide 2" />
    </office:presentation>
  </office:body>
</office:document-content>
"#;
        let zip_data = build_odf_zip(mimetype, content_xml, None);
        let parser = DocumentParser::new();
        let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
        assert_eq!(parsed.format, DocumentFormat::OdfPresentation);
        let doc = parsed.document().unwrap();
        assert_eq!(doc.content.len(), 2);
    }

    #[test]
    fn test_parse_odp_shapes_and_notes() {
        let mimetype = "application/vnd.oasis.opendocument.presentation";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1" presentation:transition-type="fade" presentation:transition-speed="fast">
        <draw:frame draw:name="Title">
          <draw:text-box>
            <text:p>Hello ODP</text:p>
          </draw:text-box>
        </draw:frame>
        <draw:frame draw:name="Image1">
          <draw:image xlink:href="Pictures/img1.png" />
        </draw:frame>
        <draw:frame draw:name="Chart1">
          <chart:chart />
        </draw:frame>
        <presentation:notes>
          <text:p>Speaker note</text:p>
        </presentation:notes>
      </draw:page>
    </office:presentation>
  </office:body>
</office:document-content>
"#;
        let zip_data = build_odf_zip(mimetype, content_xml, None);
        let parser = DocumentParser::new();
        let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

        let mut shape_count = 0;
        let mut slide_notes = 0;
        let mut transition_count = 0;
        for node in parsed.store.values() {
            if let IRNode::Slide(slide) = node {
                if slide.notes.is_some() {
                    slide_notes += 1;
                }
                if slide.transition.is_some() {
                    transition_count += 1;
                }
            }
            if let IRNode::Shape(_) = node {
                shape_count += 1;
            }
        }

        assert_eq!(shape_count, 3);
        assert_eq!(slide_notes, 1);
        assert_eq!(transition_count, 1);
    }

    #[test]
    fn test_parse_odf_security_indicators() {
        let mimetype = "application/vnd.oasis.opendocument.text";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <text:p>
        <text:a xlink:href="https://example.com">Link</text:a>
      </text:p>
      <draw:object-ole xlink:href="https://example.com/ole.bin" />
    </office:text>
  </office:body>
  <office:scripts>
    <script:script xlink:href="Scripts/macro.py" />
  </office:scripts>
</office:document-content>
"#;
        let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Scripts/macro.py" manifest:media-type="application/vnd.sun.star.script"/>
  <manifest:file-entry manifest:full-path="Object 1" manifest:media-type="application/vnd.sun.star.oleobject"/>
  <manifest:file-entry manifest:full-path="Encrypted" manifest:media-type="application/vnd.oasis.opendocument.encrypted"/>
</manifest:manifest>
"#;
        let signatures_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<ds:Signatures xmlns:ds="http://www.w3.org/2000/09/xmldsig#">
  <ds:Signature>
    <ds:SignedInfo>
      <ds:SignatureMethod Algorithm="http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"/>
      <ds:Reference>
        <ds:DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>
      </ds:Reference>
    </ds:SignedInfo>
    <ds:KeyInfo>
      <ds:X509Data>
        <ds:X509SubjectName>CN=Tester</ds:X509SubjectName>
      </ds:X509Data>
    </ds:KeyInfo>
  </ds:Signature>
</ds:Signatures>
"#;
        let zip_data = build_odf_zip_custom(
            mimetype,
            content_xml,
            None,
            Some(manifest_xml),
            vec![
                ("META-INF/documentsignatures.xml", signatures_xml),
                ("Scripts/macro.py", b"print('hi')"),
                ("Object 1", b"oledata"),
            ],
        );
        let parser = DocumentParser::new();
        let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
        docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
        let doc = parsed.document().unwrap();

        assert!(doc.security.macro_project.is_some());
        assert!(!doc.security.external_refs.is_empty());
        assert!(!doc.security.ole_objects.is_empty());

        let mut sig_count = 0;
        let mut encryption_diag = false;
        for node in parsed.store.values() {
            if let IRNode::DigitalSignature(_) = node {
                sig_count += 1;
            }
            if let IRNode::Diagnostics(diag) = node {
                if diag.entries.iter().any(|e| e.code == "ODF_ENCRYPTION") {
                    encryption_diag = true;
                }
            }
        }
        assert_eq!(sig_count, 1);
        assert!(encryption_diag);
    }

    #[test]
    fn test_parse_odt_rich_content() {
        let mimetype = "application/vnd.oasis.opendocument.text";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:dc="http://purl.org/dc/elements/1.1/">
  <office:body>
    <office:text>
      <text:p>Intro</text:p>
      <text:list text:style-name="L1">
        <text:list-item>
          <text:p>Item 1</text:p>
        </text:list-item>
      </text:list>
      <table:table>
        <table:table-row>
          <table:table-cell table:number-columns-spanned="2">
            <text:p>Cell A</text:p>
          </table:table-cell>
        </table:table-row>
      </table:table>
      <office:annotation>
        <dc:creator>Alice</dc:creator>
        <dc:date>2024-01-01</dc:date>
        <text:p>Comment body</text:p>
      </office:annotation>
      <text:note text:note-class="footnote">
        <text:note-body>
          <text:p>Footnote body</text:p>
        </text:note-body>
      </text:note>
      <text:bookmark-start text:name="bm1" />
      <text:bookmark-end text:name="bm1" />
      <text:date />
      <draw:frame draw:name="Image1">
        <draw:image xlink:href="Pictures/image1.png" />
      </draw:frame>
      <text:tracked-changes>
        <text:changed-region>
          <text:change-info>
            <dc:creator>Bob</dc:creator>
            <dc:date>2024-01-02</dc:date>
          </text:change-info>
          <text:insertion>
            <text:p>Inserted text</text:p>
          </text:insertion>
        </text:changed-region>
      </text:tracked-changes>
    </office:text>
  </office:body>
</office:document-content>
"#;
        let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
  <office:styles>
    <style:style style:name="P1" style:family="paragraph" />
  </office:styles>
</office:document-styles>
"#;
        let zip_data = build_odf_zip(mimetype, content_xml, Some(styles_xml));
        let parser = DocumentParser::new();
        let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

        let mut table_count = 0;
        let mut comment_count = 0;
        let mut footnote_count = 0;
        let mut bookmark_count = 0;
        let mut field_count = 0;
        let mut shape_count = 0;
        let mut revision_count = 0;
        let mut styles_count = 0;

        for node in parsed.store.values() {
            match node {
                IRNode::Table(_) => table_count += 1,
                IRNode::Comment(_) => comment_count += 1,
                IRNode::Footnote(_) => footnote_count += 1,
                IRNode::BookmarkStart(_) | IRNode::BookmarkEnd(_) => bookmark_count += 1,
                IRNode::Field(_) => field_count += 1,
                IRNode::Shape(_) => shape_count += 1,
                IRNode::Revision(_) => revision_count += 1,
                IRNode::StyleSet(_) => styles_count += 1,
                _ => {}
            }
        }

        assert_eq!(table_count, 1);
        assert_eq!(comment_count, 1);
        assert_eq!(footnote_count, 1);
        assert!(bookmark_count >= 2);
        assert_eq!(field_count, 1);
        assert_eq!(shape_count, 1);
        assert_eq!(revision_count, 1);
        assert_eq!(styles_count, 1);
    }

    #[test]
    fn test_odf_formula_dde_and_links() {
        let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell table:formula="of:=DDE(&quot;soffice&quot;;&quot;file:///tmp/test.ods&quot;;&quot;A1&quot;)" />
          <table:table-cell table:formula="of:=HYPERLINK(&quot;https://example.com&quot;;&quot;Example&quot;)" />
          <table:table-cell table:formula="of:=WEBSERVICE(&quot;https://example.com/api&quot;)" />
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
        let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#;
        let zip_data =
            build_odf_zip_custom(mimetype, content_xml, None, Some(manifest_xml), Vec::new());
        let parser = DocumentParser::new();
        let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
        docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
        let doc = parsed.document().unwrap();

        assert!(!doc.security.dde_fields.is_empty());
        assert!(doc
            .security
            .threat_indicators
            .iter()
            .any(|i| i.indicator_type == ThreatIndicatorType::DdeCommand));

        let mut has_formula_link = false;
        let mut has_unsupported = false;
        for node in parsed.store.values() {
            if let IRNode::ExternalReference(ext) = node {
                if ext.target.contains("example.com") {
                    has_formula_link = true;
                }
            }
            if let IRNode::Diagnostics(diag) = node {
                if diag
                    .entries
                    .iter()
                    .any(|e| e.code == "ODF_FORMULA_UNSUPPORTED_FUNCTION")
                {
                    has_unsupported = true;
                }
            }
        }
        assert!(has_formula_link);
        assert!(has_unsupported);
    }

    #[test]
    fn test_odf_encryption_metadata_diagnostics() {
        let mimetype = "application/vnd.oasis.opendocument.text";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:text><text:p>Encrypted</text:p></office:text></office:body>
</office:document-content>
"#;
        let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml">
    <manifest:encryption-data manifest:checksum-type="SHA1" manifest:checksum="YWJjZA==">
      <manifest:algorithm manifest:algorithm-name="http://www.w3.org/2001/04/xmlenc#aes256-cbc"
        manifest:initialisation-vector="MTIzNDU2Nzg5MA==" manifest:key-size="32"/>
      <manifest:key-derivation manifest:key-derivation-name="PBKDF2"
        manifest:salt="c2FsdA==" manifest:iteration-count="2048"/>
    </manifest:encryption-data>
  </manifest:file-entry>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#;
        let zip_data =
            build_odf_zip_custom(mimetype, content_xml, None, Some(manifest_xml), Vec::new());
        let parser = DocumentParser::new();
        let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

        let mut has_meta = false;
        for node in parsed.store.values() {
            if let IRNode::Diagnostics(diag) = node {
                if diag.entries.iter().any(|e| e.code == "ODF_ENCRYPTION_META") {
                    has_meta = true;
                }
            }
        }
        assert!(has_meta);
    }

    #[test]
    fn test_odf_manifest_inventory_and_parts() {
        let mimetype = "application/vnd.oasis.opendocument.text";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
  <office:body>
    <office:text>
      <text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">Hello</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
        let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Thumbnails/thumbnail.png" manifest:media-type="image/png"/>
</manifest:manifest>
"#;
        let zip_data = build_odf_zip_custom(
            mimetype,
            content_xml,
            Some(r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#),
            Some(manifest_xml),
            vec![
                (
                    "settings.xml",
                    br#"<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
                ),
                ("Thumbnails/thumbnail.png", b"pngdata"),
            ],
        );
        let parser = DocumentParser::new();
        let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

        let mut part_paths = Vec::new();
        let mut asset_paths = Vec::new();
        let mut odf_parts = Vec::new();
        for node in parsed.store.values() {
            match node {
                IRNode::ExtensionPart(part) => part_paths.push(part.path.clone()),
                IRNode::MediaAsset(asset) => asset_paths.push(asset.path.clone()),
                IRNode::Diagnostics(diag) => {
                    for entry in &diag.entries {
                        if entry.code == "ODF_PART" {
                            if let Some(path) = entry.path.as_ref() {
                                odf_parts.push(path.clone());
                            }
                        }
                    }
                }
                _ => {}
            }
        }

        assert!(part_paths.contains(&"content.xml".to_string()));
        assert!(part_paths.contains(&"styles.xml".to_string()));
        assert!(part_paths.contains(&"settings.xml".to_string()));
        assert!(asset_paths.contains(&"Thumbnails/thumbnail.png".to_string()));
        assert!(odf_parts.contains(&"content.xml".to_string()));
        assert!(odf_parts.contains(&"styles.xml".to_string()));
    }

    #[test]
    fn test_odt_headers_and_footers_from_styles() {
        let mimetype = "application/vnd.oasis.opendocument.text";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:p>Body</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
        let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:master-styles>
    <style:master-page style:name="Standard">
      <style:header>
        <text:p>Header text</text:p>
      </style:header>
      <style:footer>
        <text:p>Footer text</text:p>
      </style:footer>
    </style:master-page>
  </office:master-styles>
</office:document-styles>
"#;
        let zip_data = build_odf_zip(mimetype, content_xml, Some(styles_xml));
        let parser = DocumentParser::new();
        let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

        let mut header_texts = Vec::new();
        let mut footer_texts = Vec::new();
        for node in parsed.store.values() {
            match node {
                IRNode::Header(header) => {
                    for id in &header.content {
                        if let Some(IRNode::Paragraph(p)) = parsed.store.get(*id) {
                            for run_id in &p.runs {
                                if let Some(IRNode::Run(run)) = parsed.store.get(*run_id) {
                                    header_texts.push(run.text.clone());
                                }
                            }
                        }
                    }
                }
                IRNode::Footer(footer) => {
                    for id in &footer.content {
                        if let Some(IRNode::Paragraph(p)) = parsed.store.get(*id) {
                            for run_id in &p.runs {
                                if let Some(IRNode::Run(run)) = parsed.store.get(*run_id) {
                                    footer_texts.push(run.text.clone());
                                }
                            }
                        }
                    }
                }
                _ => {}
            }
        }

        assert!(header_texts.iter().any(|t| t == "Header text"));
        assert!(footer_texts.iter().any(|t| t == "Footer text"));
    }

    #[test]
    fn test_parse_ods_named_ranges_and_pivots() {
        let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:spreadsheet>
      <table:named-expressions>
        <table:named-range table:name="RANGE1" table:cell-range-address="Sheet1.A1:Sheet1.B2"/>
        <table:named-expression table:name="EXPR1" table:expression="of:=SUM([.A1];[.B1])"/>
      </table:named-expressions>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell table:cell-value-type="float" table:cell-value="1"/>
          <table:table-cell table:cell-value-type="float" table:cell-value="2"/>
        </table:table-row>
      </table:table>
      <table:data-pilot-table table:name="Pivot1"
        table:source-range-address="Sheet1.A1:Sheet1.B2"
        table:target-range-address="Sheet1.D1:Sheet1.E2">
        <table:data-pilot-field table:source-field-name="Field1"/>
      </table:data-pilot-table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
        let zip_data = build_odf_zip(mimetype, content_xml, None);
        let parser = DocumentParser::new();
        let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
        docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
        let doc = parsed.document().unwrap();

        assert!(!doc.defined_names.is_empty());

        let mut pivot_tables = 0;
        let mut pivot_caches = 0;
        let mut pivot_records = 0;
        for node in parsed.store.values() {
            match node {
                IRNode::PivotTable(_) => pivot_tables += 1,
                IRNode::PivotCache(_) => pivot_caches += 1,
                IRNode::PivotCacheRecords(_) => pivot_records += 1,
                _ => {}
            }
        }

        assert!(pivot_tables >= 1);
        assert!(pivot_caches >= 1);
        assert!(pivot_records >= 1);
    }

    #[test]
    fn test_parse_odp_master_pages_and_transitions() {
        let mimetype = "application/vnd.oasis.opendocument.presentation";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1" presentation:transition-type="fade" presentation:transition-speed="fast"/>
    </office:presentation>
  </office:body>
</office:document-content>
"#;
        let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
  <office:master-styles>
    <style:master-page style:name="Master1"/>
  </office:master-styles>
</office:document-styles>
"#;
        let zip_data = build_odf_zip(mimetype, content_xml, Some(styles_xml));
        let parser = DocumentParser::new();
        let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

        let mut slide_with_transition = 0;
        let mut master_page_diag = false;
        for node in parsed.store.values() {
            if let IRNode::Slide(slide) = node {
                if slide.transition.is_some() {
                    slide_with_transition += 1;
                }
            }
            if let IRNode::Diagnostics(diag) = node {
                if diag.entries.iter().any(|e| e.code == "ODF_MASTER_PAGE") {
                    master_page_diag = true;
                }
            }
        }

        assert_eq!(slide_with_transition, 1);
        assert!(master_page_diag);
    }

    #[test]
    fn test_odf_security_threat_indicators() {
        let mimetype = "application/vnd.oasis.opendocument.text";
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <text:p>
        <text:a xlink:href="https://example.com">Link</text:a>
      </text:p>
      <draw:object-ole xlink:href="https://example.com/ole.bin" />
    </office:text>
  </office:body>
  <office:scripts>
    <script:script xlink:href="Scripts/macro.py" />
  </office:scripts>
</office:document-content>
"#;
        let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Scripts/macro.py" manifest:media-type="application/vnd.sun.star.script"/>
  <manifest:file-entry manifest:full-path="Object 1" manifest:media-type="application/vnd.sun.star.oleobject"/>
</manifest:manifest>
"#;
        let zip_data = build_odf_zip_custom(
            mimetype,
            content_xml,
            None,
            Some(manifest_xml),
            vec![
                ("Scripts/macro.py", b"print('hi')"),
                ("Object 1", b"oledata"),
            ],
        );
        let parser = DocumentParser::new();
        let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
        docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
        let doc = parsed.document().unwrap();

        assert!(doc.security.macro_project.is_some());
        assert!(!doc.security.external_refs.is_empty());
        assert!(!doc.security.ole_objects.is_empty());
        assert!(doc
            .security
            .threat_indicators
            .iter()
            .any(|i| i.indicator_type == ThreatIndicatorType::RemoteResource));
        assert!(doc
            .security
            .threat_indicators
            .iter()
            .any(|i| i.indicator_type == ThreatIndicatorType::OleObject));
    }
}
