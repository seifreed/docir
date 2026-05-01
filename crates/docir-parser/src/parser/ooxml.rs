use super::{
    read_all_with_limit, rel_type, run_parser_pipeline, security, ContentTypes, Cursor,
    DocumentFormat, FormatParser, IRNode, IrStore, NodeId, NormalizeStage, PackageReader,
    ParseError, ParseMetrics, ParseStage, ParsedDocument, ParserConfig, PostprocessStage, Read,
    Relationships, SecureZipReader, Seek, SeekFrom,
};
#[path = "ooxml_docx_parts.rs"]
mod ooxml_docx_parts;

pub(crate) struct DocxWordParts {
    pub(super) styles_id: Option<NodeId>,
    pub(super) styles_with_effects_id: Option<NodeId>,
    pub(super) numbering_id: Option<NodeId>,
    pub(super) comments: Vec<NodeId>,
    pub(super) footnotes: Vec<NodeId>,
    pub(super) endnotes: Vec<NodeId>,
    pub(super) settings_id: Option<NodeId>,
    pub(super) web_settings_id: Option<NodeId>,
    pub(super) font_table_id: Option<NodeId>,
    pub(super) comments_ext_id: Option<NodeId>,
    pub(super) comments_id_map_id: Option<NodeId>,
    pub(super) glossary_id: Option<NodeId>,
}

type DocxAnnotationParts = (
    Vec<NodeId>,
    Vec<NodeId>,
    Vec<NodeId>,
    Option<NodeId>,
    Option<NodeId>,
);
struct HeaderFooterSpec {
    rel_type: &'static str,
    kind: crate::ooxml::docx::document::HeaderFooterKind,
}

/// Main OOXML parser.
pub struct OoxmlParser {
    pub(super) config: ParserConfig,
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
    pub fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        run_parser_pipeline(self, reader)
    }
}

impl ParseStage for OoxmlParser {
    fn parse_stage<R: Read + Seek>(&self, mut reader: R) -> Result<ParsedDocument, ParseError> {
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
}

impl NormalizeStage for OoxmlParser {}

impl PostprocessStage for OoxmlParser {}

impl OoxmlParser {
    fn init_metrics(&self) -> Option<ParseMetrics> {
        if self.config.enable_metrics {
            Some(ParseMetrics::default())
        } else {
            None
        }
    }

    fn read_content_types(
        &self,
        zip: &mut impl PackageReader,
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

    fn read_package_relationships(
        &self,
        zip: &mut impl PackageReader,
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

    fn parse_main(
        &self,
        zip: &mut impl PackageReader,
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

    fn validate_structure(
        &self,
        zip: &mut impl PackageReader,
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

    pub(super) fn finalize_ooxml_document(
        &self,
        zip: &mut impl PackageReader,
        content_types: &ContentTypes,
        store: &mut IrStore,
        root_id: NodeId,
        metrics: &mut Option<ParseMetrics>,
    ) -> Result<(), ParseError> {
        let metadata_id = self.parse_metadata(zip)?;
        if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
            doc.metadata = metadata_id;
        }
        if metadata_id.is_some() {
            if let Some(metadata) = self.build_metadata(zip) {
                store.insert(IRNode::Metadata(metadata));
            }
        }

        let start = std::time::Instant::now();
        self.parse_shared_parts(zip, content_types, store, root_id)?;
        if let Some(m) = metrics.as_mut() {
            m.shared_parts_ms = start.elapsed().as_millis();
        }

        if self.config.scan_security_on_parse {
            let start = std::time::Instant::now();
            let scanner = security::SecurityScanner::new(&self.config);
            scanner.scan_zip(zip, store)?;
            if let Some(m) = metrics.as_mut() {
                m.security_scan_ms = start.elapsed().as_millis();
            }
        }

        self.post_process_ooxml(zip, content_types, store, root_id, metrics)?;
        Ok(())
    }

    pub(super) fn build_parsed_document(
        &self,
        root_id: NodeId,
        format: DocumentFormat,
        store: IrStore,
    ) -> ParsedDocument {
        ParsedDocument {
            root_id,
            format,
            store,
            metrics: None,
        }
    }
}
