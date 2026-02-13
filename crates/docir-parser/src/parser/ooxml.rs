use super::*;

pub(super) struct DocxWordParts {
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

    pub(super) fn parse_docx_word_parts(
        &self,
        zip: &mut impl PackageReader,
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

    pub(super) fn parse_docx_headers_footers(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> Result<HashMap<String, NodeId>, ParseError> {
        let mut map = HashMap::new();

        self.parse_docx_header_footer_kind(
            zip,
            main_part_path,
            doc_rels,
            parser,
            rel_type::HEADER,
            crate::ooxml::docx::document::HeaderFooterKind::Header,
            &mut map,
        )?;
        self.parse_docx_header_footer_kind(
            zip,
            main_part_path,
            doc_rels,
            parser,
            rel_type::FOOTER,
            crate::ooxml::docx::document::HeaderFooterKind::Footer,
            &mut map,
        )?;

        Ok(map)
    }

    fn parse_docx_header_footer_kind(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
        rel_type: &str,
        kind: crate::ooxml::docx::document::HeaderFooterKind,
        map: &mut HashMap<String, NodeId>,
    ) -> Result<(), ParseError> {
        for rel in doc_rels.get_by_type(rel_type) {
            let part_path = Relationships::resolve_target(main_part_path, &rel.target);
            let rels = read_relationships_optional(zip, &part_path);
            let xml = zip.read_file_string(&part_path)?;
            let node_id = parser.parse_header_footer(&xml, &part_path, kind, &rels)?;
            map.insert(rel.id.clone(), node_id);
        }

        Ok(())
    }

    fn parse_docx_comments(
        &self,
        zip: &mut impl PackageReader,
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

    fn parse_docx_notes(
        &self,
        zip: &mut impl PackageReader,
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

    fn parse_docx_font_table(
        &self,
        zip: &mut impl PackageReader,
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

    fn parse_docx_part_by_rel<F>(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        rel_type: &str,
        mut parse: F,
    ) -> Option<NodeId>
    where
        F: FnMut(&str, &str) -> Option<NodeId>,
    {
        let (part_path, xml) =
            self.read_xml_part_by_rel_optional(zip, main_part_path, doc_rels, rel_type)?;
        parse(&part_path, &xml)
    }

    fn parse_docx_part_by_rel_with_span<F, S>(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        rel_type: &str,
        parser: &mut DocxParser,
        parse: F,
        set_span: S,
    ) -> Option<NodeId>
    where
        F: FnOnce(&mut DocxParser, &str, &str) -> Option<NodeId>,
        S: FnOnce(&mut IrStore, NodeId, &str),
    {
        let (part_path, xml) =
            self.read_xml_part_by_rel_optional(zip, main_part_path, doc_rels, rel_type)?;
        let id = parse(parser, &part_path, &xml)?;
        set_span(parser.store_mut(), id, &part_path);
        Some(id)
    }

    fn parse_docx_part_by_path<F>(
        &self,
        zip: &mut impl PackageReader,
        part_path: &str,
        mut parse: F,
    ) -> Option<NodeId>
    where
        F: FnMut(&str, &str) -> Option<NodeId>,
    {
        let xml = self.read_xml_part_optional(zip, part_path)?;
        parse(part_path, &xml)
    }

    fn parse_docx_part_by_path_with_span<F, S>(
        &self,
        zip: &mut impl PackageReader,
        part_path: &str,
        parser: &mut DocxParser,
        parse: F,
        set_span: S,
    ) -> Option<NodeId>
    where
        F: FnOnce(&mut DocxParser, &str, &str) -> Option<NodeId>,
        S: FnOnce(&mut IrStore, NodeId, &str),
    {
        let xml = self.read_xml_part_optional(zip, part_path)?;
        let id = parse(parser, part_path, &xml)?;
        set_span(parser.store_mut(), id, part_path);
        Some(id)
    }

    fn read_xml_part_by_rel_optional(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        rel_type: &str,
    ) -> Option<(String, String)> {
        read_xml_part_by_rel(zip, main_part_path, doc_rels, rel_type)
            .ok()
            .flatten()
    }

    fn read_xml_part_optional(
        &self,
        zip: &mut impl PackageReader,
        part_path: &str,
    ) -> Option<String> {
        read_xml_part(zip, part_path).ok().flatten()
    }
}
