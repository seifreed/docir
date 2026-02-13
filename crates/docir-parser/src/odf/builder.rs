use super::*;
use crate::parse_utils::{finalize_document, init_document_state};

impl OdfParser {
    /// Parses from any reader.
    pub fn parse_reader<R: Read + Seek>(
        &self,
        mut reader: R,
    ) -> Result<ParsedDocument, ParseError> {
        enforce_input_size(&mut reader, self.config.max_input_size)?;
        let mut zip = SecureZipReader::new(reader, self.config.zip_config.clone())?;

        let (format, manifest_entries) = self.load_mimetype_and_manifest(&mut zip)?;

        if !zip.contains("content.xml") {
            return Err(ParseError::MissingPart("content.xml".to_string()));
        }

        let (mut store, mut doc, mut diagnostics) = init_document_state(format);

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

        let (styles_xml, settings_xml, signatures_xml) =
            self.load_styles_settings_signatures(&mut zip, &mut store, &mut doc)?;

        self.emit_fast_mode_diagnostics(
            &mut diagnostics,
            fast_mode,
            content_size,
            content_xml.is_none(),
        );
        let manifest_index = collect_manifest_index(&manifest_entries, &mut diagnostics);
        let file_names = collect_shared_parts(&mut zip, &manifest_index, &mut store, &mut doc);

        self.process_odf_styles_and_layouts(
            &self.config,
            styles_xml.as_deref(),
            content_xml.as_deref(),
            fast_mode,
            &mut store,
            &mut doc,
            &mut diagnostics,
        )?;

        self.capture_fast_mode_spreadsheet_chunks(
            fast_mode,
            format,
            content_bytes.as_deref(),
            &mut store,
            &mut doc,
            &mut diagnostics,
        );

        let mut macro_project = build_odf_macro_project(
            &manifest_entries,
            &content_xml,
            &styles_xml,
            &settings_xml,
            &file_names,
            &mut store,
        );
        self.insert_macro_project(&mut store, &mut macro_project);

        let scanner = DefaultSecurityScanner;
        scanner.scan_odf(
            content_xml.as_deref(),
            styles_xml.as_deref(),
            settings_xml.as_deref(),
            &file_names,
            &mut zip,
            &mut store,
            &mut doc,
            &mut diagnostics,
        );

        self.insert_signatures(signatures_xml.as_deref(), &mut store, &mut doc);

        self.emit_encryption_diagnostics(&manifest_entries, &mut diagnostics);

        attach_diagnostics_if_any(&mut store, &mut doc, diagnostics);

        self.add_filter_diagnostics(content_xml.as_deref(), &mut store, &mut doc);

        self.add_defined_names(format, content_bytes.as_deref(), &mut store, &mut doc);

        Ok(finalize_document(format, store, doc))
    }

    fn capture_fast_mode_spreadsheet_chunks(
        &self,
        fast_mode: bool,
        format: DocumentFormat,
        content_bytes: Option<&[u8]>,
        store: &mut IrStore,
        doc: &mut Document,
        diagnostics: &mut Diagnostics,
    ) {
        if !(fast_mode && format == DocumentFormat::OdfSpreadsheet) {
            return;
        }
        let Some(bytes) = content_bytes else {
            return;
        };
        let chunks = spreadsheet_chunks::extract_spreadsheet_table_chunks(bytes);
        for (idx, chunk) in chunks.iter().enumerate() {
            let sheet_name = spreadsheet_chunks::table_name_from_chunk(&chunk.bytes, (idx + 1) as u32);
            let path = format!(
                "content.xml#sheet:{}@{}-{}",
                sheet_name, chunk.start, chunk.end
            );
            let mut part = ExtensionPart::new(
                path,
                (chunk.end.saturating_sub(chunk.start) + 1) as u64,
                ExtensionPartKind::VendorSpecific,
            );
            part.content_type = Some("application/vnd.docir.odf.lazy-sheet+xml".to_string());
            part.span = Some(SourceSpan::new("content.xml"));
            let part_id = part.id;
            store.insert(IRNode::ExtensionPart(part));
            doc.shared_parts.push(part_id);
            push_info(
                diagnostics,
                "ODF_LAZY_SHEET",
                format!(
                    "Lazy sheet range stored for {} ({}-{})",
                    sheet_name, chunk.start, chunk.end
                ),
                Some("content.xml"),
            );
        }
    }

    fn insert_macro_project(
        &self,
        store: &mut IrStore,
        macro_project: &mut Option<docir_core::security::MacroProject>,
    ) {
        if let Some(project) = macro_project.take() {
            store.insert(IRNode::MacroProject(project));
        }
    }

    fn insert_signatures(
        &self,
        signatures_xml: Option<&str>,
        store: &mut IrStore,
        doc: &mut Document,
    ) {
        let Some(sig_xml) = signatures_xml else {
            return;
        };
        let sigs = parse_odf_signatures(sig_xml);
        for sig in sigs {
            let sig_id = sig.id;
            store.insert(IRNode::DigitalSignature(sig));
            doc.shared_parts.push(sig_id);
        }
    }

    fn emit_encryption_diagnostics(
        &self,
        manifest_entries: &[OdfManifestEntry],
        diagnostics: &mut Diagnostics,
    ) {
        let encrypted_entries = encrypted_manifest_entries(manifest_entries);
        for entry in manifest_entries {
            if let Some(message) = format_odf_encryption_metadata(entry) {
                push_info(
                    diagnostics,
                    "ODF_ENCRYPTION_META",
                    message,
                    Some("META-INF/manifest.xml"),
                );
            }
        }
        if encrypted_entries.is_empty() {
            return;
        }
        push_warning(
            diagnostics,
            "ODF_ENCRYPTION",
            "ODF encrypted content detected in manifest".to_string(),
            Some("META-INF/manifest.xml"),
        );
        for path in encrypted_entries {
            push_warning(
                diagnostics,
                "ODF_ENCRYPTED_PART",
                format!("Encrypted ODF part not decrypted: {}", path),
                Some("META-INF/manifest.xml"),
            );
        }
    }

    fn add_filter_diagnostics(
        &self,
        content_xml: Option<&str>,
        store: &mut IrStore,
        doc: &mut Document,
    ) {
        let Some(xml) = content_xml else {
            return;
        };
        for filter in scan_odf_filters(xml) {
            let mut diag = Diagnostics::new();
            push_entry(
                &mut diag.entries,
                DiagnosticSeverity::Info,
                "ODF_FILTER",
                format!("ODF filter detected: {}", filter),
                Some("content.xml"),
            );
            let diag_id = diag.id;
            store.insert(IRNode::Diagnostics(diag));
            doc.add_diagnostic(diag_id);
        }
    }

    fn add_defined_names(
        &self,
        format: DocumentFormat,
        content_bytes: Option<&[u8]>,
        store: &mut IrStore,
        doc: &mut Document,
    ) {
        if format != DocumentFormat::OdfSpreadsheet {
            return;
        }
        let Some(bytes) = content_bytes else {
            return;
        };
        for name in parse_ods_named_ranges(bytes) {
            let id = name.id;
            store.insert(IRNode::DefinedName(name));
            doc.defined_names.push(id);
        }
    }

    fn emit_fast_mode_diagnostics(
        &self,
        diagnostics: &mut Diagnostics,
        fast_mode: bool,
        content_size: Option<u64>,
        skipped_scans: bool,
    ) {
        if !fast_mode {
            return;
        }

        let size = content_size.unwrap_or(0);
        push_info(
            diagnostics,
            "ODF_FAST_MODE",
            format!(
                "ODF fast mode enabled (content.xml: {} bytes, threshold: {} bytes, sample_rows: {}, sample_cols: {})",
                size,
                self.config.odf.fast_threshold_bytes,
                self.config.odf.fast_sample_rows,
                self.config.odf.fast_sample_cols
            ),
            Some("content.xml"),
        );
        if skipped_scans {
            push_warning(
                diagnostics,
                "ODF_FAST_SKIP_SCAN",
                "Fast mode skipped content.xml security scans to reduce processing time"
                    .to_string(),
                Some("content.xml"),
            );
        }
    }

    fn process_odf_styles_and_layouts(
        &self,
        config: &ParserConfig,
        styles_xml: Option<&str>,
        content_xml: Option<&str>,
        fast_mode: bool,
        store: &mut IrStore,
        doc: &mut Document,
        diagnostics: &mut Diagnostics,
    ) -> Result<(), ParseError> {
        if let Some(xml) = styles_xml {
            let masters = parse_master_pages(xml);
            for name in masters {
                push_info(
                    diagnostics,
                    "ODF_MASTER_PAGE",
                    format!("ODF master page detected: {}", name),
                    Some("styles.xml"),
                );
            }
            let layouts = parse_page_layouts(xml);
            for name in layouts {
                push_info(
                    diagnostics,
                    "ODF_PAGE_LAYOUT",
                    format!("ODF page layout detected: {}", name),
                    Some("styles.xml"),
                );
            }
            let (headers, footers) = parse_odf_headers_footers(xml, store, config)?;
            for header_id in headers {
                doc.shared_parts.push(header_id);
            }
            for footer_id in footers {
                doc.shared_parts.push(footer_id);
            }
        }

        if !fast_mode {
            if let Some(xml) = content_xml {
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
                    push_info(
                        diagnostics,
                        "ODF_MASTER_PAGE",
                        format!("ODF master page detected: {}", name),
                        Some("content.xml"),
                    );
                }
                let layouts = parse_page_layouts(xml);
                for name in layouts {
                    push_info(
                        diagnostics,
                        "ODF_PAGE_LAYOUT",
                        format!("ODF page layout detected: {}", name),
                        Some("content.xml"),
                    );
                }
            }
        }

        Ok(())
    }
}
