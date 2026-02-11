use super::*;

impl OoxmlParser {
    /// Parse a DOCX document.
    pub(super) fn parse_docx<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        main_part_path: &str,
        content_types: &ContentTypes,
        metrics: &mut Option<ParseMetrics>,
    ) -> Result<ParsedDocument, ParseError> {
        // Read main document
        let document_xml = zip.read_file_string(main_part_path)?;

        // Get document relationships
        let rels_path = Self::get_rels_path(main_part_path);
        let doc_rels = if zip.contains(&rels_path) {
            let rels_xml = zip.read_file_string(&rels_path)?;
            Relationships::parse(&rels_xml)?
        } else {
            Relationships::default()
        };

        // Parse document
        let mut docx_parser = DocxParser::new();

        let header_footer_map =
            self.parse_docx_headers_footers(zip, main_part_path, &doc_rels, &mut docx_parser)?;

        let (
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
        ) = self.parse_docx_word_parts(zip, main_part_path, &doc_rels, &mut docx_parser);

        let root_id =
            docx_parser.parse_document(&document_xml, &doc_rels, Some(&header_footer_map))?;
        let mut store = docx_parser.into_store();

        if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
            doc.styles = styles_id;
            doc.styles_with_effects = styles_with_effects_id;
            doc.numbering = numbering_id;
            doc.comments = comments;
            doc.footnotes = footnotes;
            doc.endnotes = endnotes;
            doc.settings = settings_id;
            doc.web_settings = web_settings_id;
            doc.font_table = font_table_id;
            doc.comments_extended = comments_ext_id;
            doc.comment_id_map = comments_id_map_id;
            if let Some(glossary_id) = glossary_id {
                doc.shared_parts.push(glossary_id);
            }
        }

        // Parse metadata
        let metadata_id = self.parse_metadata(zip)?;
        if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
            doc.metadata = metadata_id;
        }
        if metadata_id.is_some() {
            if let Some(metadata) = self.build_metadata(zip) {
                store.insert(IRNode::Metadata(metadata));
            }
        }

        // Shared parts
        let start = std::time::Instant::now();
        self.parse_shared_parts(zip, content_types, &mut store, root_id)?;
        if let Some(m) = metrics.as_mut() {
            m.shared_parts_ms = start.elapsed().as_millis();
        }

        // Scan for security content
        let start = std::time::Instant::now();
        self.scan_security_content(zip, &mut store, root_id, content_types)?;
        if let Some(m) = metrics.as_mut() {
            m.security_scan_ms = start.elapsed().as_millis();
        }

        // Link shapes/animations to shared parts (charts, SmartArt, media, OLE)
        self.link_shapes_to_shared_parts(&mut store);

        // Extension parts + diagnostics
        let start = std::time::Instant::now();
        self.add_extension_parts_and_diagnostics(zip, content_types, &mut store, root_id)?;
        if let Some(m) = metrics.as_mut() {
            m.extension_parts_ms = start.elapsed().as_millis();
        }

        // Deterministic normalization
        let start = std::time::Instant::now();
        normalize_store(&mut store, root_id);
        if let Some(m) = metrics.as_mut() {
            m.normalization_ms = start.elapsed().as_millis();
        }

        Ok(ParsedDocument {
            root_id,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        })
    }
}
