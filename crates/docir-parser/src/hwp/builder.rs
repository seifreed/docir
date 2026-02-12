use super::*;
use crate::parse_utils::{finalize_and_normalize, init_store_and_document};

impl HwpParser {
    pub fn parse_reader<R: Read + Seek>(
        &self,
        mut reader: R,
    ) -> Result<ParsedDocument, ParseError> {
        let data = read_all_with_limit(reader, self.config.max_input_size)?;

        let cfb = Cfb::parse(data)?;
        let (mut store, mut doc) = init_store_and_document(DocumentFormat::Hwp);

        let mut stream_names = cfb.list_streams();
        stream_names.sort();
        let mut shared_parts = Vec::new();
        let mut sections = Vec::new();
        let mut diagnostics = build_hwp_diagnostics(DocumentFormat::Hwp, &stream_names);

        let header_ctx = self.build_header_context(&cfb, &mut diagnostics)?;

        if self.config.hwp.dump_streams {
            dump_hwp_streams(
                &cfb,
                &stream_names,
                header_ctx.compressed,
                header_ctx.encrypted,
                header_ctx.hwp_password,
                header_ctx.force_parse,
                header_ctx.try_raw_encrypted,
                &mut diagnostics,
            );
        }

        for path in &stream_names {
            let size = cfb.stream_size(path).unwrap_or(0);
            if path.starts_with("BinData/") {
                let mut asset = MediaAsset::new(path, MediaType::Other, size);
                asset.span = Some(SourceSpan::new(path));
                let asset_id = asset.id;
                store.insert(IRNode::MediaAsset(asset));
                shared_parts.push(asset_id);
            } else {
                let mut part = ExtensionPart::new(path, size, ExtensionPartKind::Legacy);
                part.span = Some(SourceSpan::new(path));
                let part_id = part.id;
                store.insert(IRNode::ExtensionPart(part));
                shared_parts.push(part_id);
            }
        }

        let docinfo_data = cfb.read_stream("DocInfo");
        let mut docinfo_section_count = None;
        if let Some(data) = docinfo_data {
            let data = match prepare_hwp_stream_data(
                &data,
                header_ctx.encrypted,
                header_ctx.hwp_password,
                header_ctx.force_parse,
                header_ctx.try_raw_encrypted,
                "DocInfo",
                &mut diagnostics,
            ) {
                Some(bytes) => bytes,
                None => {
                    push_warning(
                        &mut diagnostics,
                        "HWP_DOCINFO_SKIP",
                        "DocInfo skipped due to encryption or decryption failure".to_string(),
                        Some("DocInfo"),
                    );
                    Vec::new()
                }
            };
            match maybe_decompress_stream(&data, header_ctx.compressed, "DocInfo") {
                Ok(bytes) => {
                    docinfo_section_count = parse_docinfo_section_count(&bytes)?;
                }
                Err(err) => {
                    push_warning(
                        &mut diagnostics,
                        "HWP_DECOMPRESS_FAIL",
                        err.to_string(),
                        Some("DocInfo"),
                    );
                }
            }
            if let Some(count) = docinfo_section_count {
                push_info(
                    &mut diagnostics,
                    "HWP_SECTION_COUNT",
                    format!("DocInfo section count: {}", count),
                    Some("DocInfo"),
                );
            }
        }

        if header_ctx.allow_parse {
            for path in &stream_names {
                if path.starts_with("BodyText/Section") {
                    let data = cfb
                        .read_stream(path)
                        .ok_or_else(|| ParseError::MissingPart(path.to_string()))?;
                    let data = match prepare_hwp_stream_data(
                        &data,
                        header_ctx.encrypted,
                        header_ctx.hwp_password,
                        header_ctx.force_parse,
                        header_ctx.try_raw_encrypted,
                        path,
                        &mut diagnostics,
                    ) {
                        Some(bytes) => bytes,
                        None => continue,
                    };
                    let data = match maybe_decompress_stream(&data, header_ctx.compressed, path) {
                        Ok(bytes) => bytes,
                        Err(err) => {
                            push_warning(
                                &mut diagnostics,
                                "HWP_DECOMPRESS_FAIL",
                                err.to_string(),
                                Some(path),
                            );
                            continue;
                        }
                    };
                    let paragraph_ids = parse_hwp_section_stream(&data, path, &mut store)?;
                    let mut section = Section::new();
                    section.name = Some(path.clone());
                    section.content = paragraph_ids;
                    section.span = Some(SourceSpan::new(path));
                    let section_id = section.id;
                    store.insert(IRNode::Section(section));
                    sections.push(section_id);
                }
            }
        }

        if let Some(expected) = docinfo_section_count {
            if expected as usize != sections.len() {
                push_warning(
                    &mut diagnostics,
                    "HWP_SECTION_MISMATCH",
                    format!(
                        "section count mismatch: docinfo={} parsed={}",
                        expected,
                        sections.len()
                    ),
                    Some("DocInfo"),
                );
            }
        }

        if let Some(script_data) = cfb.read_stream("Scripts/DefaultJScript") {
            if let Some(_project_id) =
                parse_default_jscript(&script_data, &mut store, "Scripts/DefaultJScript")
            {
            }
        }

        if header_ctx.allow_parse {
            let externals = scan_hwp_external_refs(
                &cfb,
                &stream_names,
                header_ctx.compressed,
                header_ctx.encrypted,
                header_ctx.hwp_password,
                header_ctx.force_parse,
                header_ctx.try_raw_encrypted,
                &mut diagnostics,
            );
            for ext in externals {
                store.insert(IRNode::ExternalReference(ext));
            }
        }

        for path in &stream_names {
            if path.starts_with("BinData/") {
                let lower = path.to_ascii_lowercase();
                if lower.contains("ole") || lower.contains("object") {
                    let mut ole = OleObject::new();
                    ole.name = Some(path.clone());
                    ole.size_bytes = cfb.stream_size(path).unwrap_or(0);
                    store.insert(IRNode::OleObject(ole));
                }
            }
        }

        doc.shared_parts = shared_parts;
        doc.content = sections;

        attach_diagnostics_if_any(&mut store, &mut doc, diagnostics);

        Ok(finalize_and_normalize(DocumentFormat::Hwp, store, doc))
    }

    fn build_header_context<'a>(
        &'a self,
        cfb: &Cfb,
        diagnostics: &mut Diagnostics,
    ) -> Result<HwpHeaderContext<'a>, ParseError> {
        let header_data = cfb
            .read_stream("FileHeader")
            .ok_or_else(|| ParseError::MissingPart("FileHeader".to_string()))?;
        let header = parse_file_header(&header_data)?;
        push_info(
            diagnostics,
            "HWP_HEADER",
            format!(
                "HWP header: version=0x{:08X} flags=0x{:08X}",
                header.version, header.flags
            ),
            Some("FileHeader"),
        );

        let compressed = header.flags & 0x01 != 0;
        let encrypted = header.flags & 0x02 != 0;
        let force_parse = self.config.hwp.force_parse_encrypted;
        let hwp_password = self.config.hwp.password.as_deref();
        let try_raw_encrypted = encrypted && hwp_password.is_none();
        let allow_parse = !encrypted || force_parse || hwp_password.is_some() || try_raw_encrypted;
        if encrypted {
            push_warning(
                diagnostics,
                "HWP_ENCRYPTED",
                "HWP file is encrypted; content parsing skipped".to_string(),
                Some("FileHeader"),
            );
            if force_parse {
                push_warning(
                    diagnostics,
                    "HWP_FORCE_PARSE",
                    "HWP force-parse enabled for encrypted file".to_string(),
                    Some("FileHeader"),
                );
            }
            if hwp_password.is_some() {
                push_info(
                    diagnostics,
                    "HWP_DECRYPT_ATTEMPT",
                    "HWP decryption attempt enabled".to_string(),
                    Some("FileHeader"),
                );
            }
            if try_raw_encrypted {
                push_warning(
                    diagnostics,
                    "HWP_ENCRYPTED_PARTIAL",
                    "HWP encrypted without password; attempting partial parse of readable streams"
                        .to_string(),
                    Some("FileHeader"),
                );
            }
        }

        Ok(HwpHeaderContext {
            compressed,
            encrypted,
            force_parse,
            hwp_password,
            try_raw_encrypted,
            allow_parse,
        })
    }
}

struct HwpHeaderContext<'a> {
    compressed: bool,
    encrypted: bool,
    force_parse: bool,
    hwp_password: Option<&'a str>,
    try_raw_encrypted: bool,
    allow_parse: bool,
}

impl HwpxParser {
    pub fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        let mut reader = reader;
        enforce_input_size(&mut reader, self.config.max_input_size)?;
        let mut zip = SecureZipReader::new(reader, self.config.zip_config.clone())?;

        let (mut store, mut doc) = init_store_and_document(DocumentFormat::Hwpx);

        let mut file_names: Vec<String> = zip.file_names().map(|s| s.to_string()).collect();
        file_names.sort();

        let mut shared_parts = Vec::new();
        let mut media_assets = Vec::new();
        let mut media_lookup: HashMap<String, NodeId> = HashMap::new();
        for path in &file_names {
            let size = zip.file_size(path).unwrap_or(0);
            if path.starts_with("BinData/") {
                let media_type = media_type_from_path(path);
                let mut asset = MediaAsset::new(path, media_type, size);
                asset.span = Some(SourceSpan::new(path));
                let asset_id = asset.id;
                store.insert(IRNode::MediaAsset(asset));
                media_assets.push(asset_id);
                media_lookup.insert(path.clone(), asset_id);
            }
            let mut part = ExtensionPart::new(path, size, ExtensionPartKind::VendorSpecific);
            part.span = Some(SourceSpan::new(path));
            let part_id = part.id;
            store.insert(IRNode::ExtensionPart(part));
            shared_parts.push(part_id);
        }
        doc.shared_parts = shared_parts;
        doc.shared_parts.extend(media_assets);

        let mut section_ids = Vec::new();
        for path in &file_names {
            if is_hwpx_section(path) {
                let xml = zip.read_file_string(path)?;
                let paragraph_ids = parse_hwpx_section(
                    &xml,
                    path,
                    &mut store,
                    &mut doc.comments,
                    &mut doc.footnotes,
                    &mut doc.endnotes,
                    &media_lookup,
                )?;
                let mut section = Section::new();
                section.name = Some(path.clone());
                section.content = paragraph_ids;
                section.span = Some(SourceSpan::new(path));
                let section_id = section.id;
                store.insert(IRNode::Section(section));
                section_ids.push(section_id);
            }
        }
        doc.content = section_ids;

        let mut header_ids = Vec::new();
        let mut footer_ids = Vec::new();
        let mut master_ids = Vec::new();

        for path in &file_names {
            if is_hwpx_header(path) {
                let xml = zip.read_file_string(path)?;
                let content = parse_hwpx_section(
                    &xml,
                    path,
                    &mut store,
                    &mut doc.comments,
                    &mut doc.footnotes,
                    &mut doc.endnotes,
                    &media_lookup,
                )?;
                let mut header = Header::new();
                header.content = content;
                header.span = Some(SourceSpan::new(path));
                let header_id = header.id;
                store.insert(IRNode::Header(header));
                header_ids.push(header_id);
            } else if is_hwpx_footer(path) {
                let xml = zip.read_file_string(path)?;
                let content = parse_hwpx_section(
                    &xml,
                    path,
                    &mut store,
                    &mut doc.comments,
                    &mut doc.footnotes,
                    &mut doc.endnotes,
                    &media_lookup,
                )?;
                let mut footer = Footer::new();
                footer.content = content;
                footer.span = Some(SourceSpan::new(path));
                let footer_id = footer.id;
                store.insert(IRNode::Footer(footer));
                footer_ids.push(footer_id);
            } else if is_hwpx_master(path) {
                let xml = zip.read_file_string(path)?;
                let content = parse_hwpx_section(
                    &xml,
                    path,
                    &mut store,
                    &mut doc.comments,
                    &mut doc.footnotes,
                    &mut doc.endnotes,
                    &media_lookup,
                )?;
                let mut section = Section::new();
                section.name = Some(format!("master:{}", path));
                section.content = content;
                section.span = Some(SourceSpan::new(path));
                let section_id = section.id;
                store.insert(IRNode::Section(section));
                master_ids.push(section_id);
            }
        }

        for section_id in &doc.content {
            if let Some(IRNode::Section(section)) = store.get_mut(*section_id) {
                section.headers.extend(header_ids.iter().copied());
                section.footers.extend(footer_ids.iter().copied());
            }
        }
        doc.shared_parts.extend(master_ids);

        if zip.contains("Contents/content.hpf") {
            if let Ok(xml) = zip.read_file_string("Contents/content.hpf") {
                if let Some(style_set) = parse_hwpx_styles(&xml, "Contents/content.hpf") {
                    let style_id = style_set.id;
                    store.insert(IRNode::StyleSet(style_set));
                    doc.styles = Some(style_id);
                }
            }
        }

        scan_hwpx_security(&file_names, &mut zip, &mut store, &mut doc);

        let diagnostics = build_hwp_diagnostics(DocumentFormat::Hwpx, &file_names);
        attach_diagnostics_if_any(&mut store, &mut doc, diagnostics);

        Ok(finalize_and_normalize(DocumentFormat::Hwpx, store, doc))
    }
}
