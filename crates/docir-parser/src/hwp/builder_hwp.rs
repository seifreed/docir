use super::super::{parse_file_header, scan_hwp_external_refs, HwpParser};
use crate::diagnostics::{push_info, push_warning};
use crate::error::ParseError;
use crate::ole::Cfb;
use docir_core::ir::{
    Diagnostics, ExtensionPart, ExtensionPartKind, IRNode, MediaAsset, MediaType, Section,
};
use docir_core::security::OleObject;
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;

use super::super::io::prepare_hwp_stream_data;
use super::super::legacy::{
    maybe_decompress_stream, parse_default_jscript, parse_docinfo_section_count,
    parse_hwp_section_stream,
};

pub(crate) struct HwpHeaderContext<'a> {
    pub(crate) compressed: bool,
    pub(crate) encrypted: bool,
    pub(crate) force_parse: bool,
    pub(crate) hwp_password: Option<&'a str>,
    pub(crate) try_raw_encrypted: bool,
    pub(crate) allow_parse: bool,
}

impl HwpParser {
    pub(super) fn collect_stream_parts(
        &self,
        cfb: &Cfb,
        stream_names: &[String],
        store: &mut IrStore,
    ) -> Vec<NodeId> {
        let mut shared_parts = Vec::new();
        for path in stream_names {
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
        shared_parts
    }

    pub(super) fn read_docinfo_section_count(
        &self,
        cfb: &Cfb,
        header_ctx: &HwpHeaderContext<'_>,
        diagnostics: &mut Diagnostics,
    ) -> Result<Option<u32>, ParseError> {
        let Some(data) = cfb.read_stream("DocInfo") else {
            return Ok(None);
        };

        let data = match prepare_hwp_stream_data(
            &data,
            header_ctx.encrypted,
            header_ctx.hwp_password,
            header_ctx.force_parse,
            header_ctx.try_raw_encrypted,
            "DocInfo",
            diagnostics,
        ) {
            Some(bytes) => bytes,
            None => {
                push_warning(
                    diagnostics,
                    "HWP_DOCINFO_SKIP",
                    "DocInfo skipped due to encryption or decryption failure".to_string(),
                    Some("DocInfo"),
                );
                Vec::new()
            }
        };

        let count = match maybe_decompress_stream(&data, header_ctx.compressed, "DocInfo") {
            Ok(bytes) => parse_docinfo_section_count(&bytes)?.map(u32::from),
            Err(err) => {
                push_warning(
                    diagnostics,
                    "HWP_DECOMPRESS_FAIL",
                    err.to_string(),
                    Some("DocInfo"),
                );
                None
            }
        };

        if let Some(count) = count {
            push_info(
                diagnostics,
                "HWP_SECTION_COUNT",
                format!("DocInfo section count: {}", count),
                Some("DocInfo"),
            );
        }
        Ok(count)
    }

    pub(super) fn parse_sections(
        &self,
        cfb: &Cfb,
        stream_names: &[String],
        header_ctx: &HwpHeaderContext<'_>,
        store: &mut IrStore,
        diagnostics: &mut Diagnostics,
    ) -> Result<Vec<NodeId>, ParseError> {
        if !header_ctx.allow_parse {
            return Ok(Vec::new());
        }

        let mut sections = Vec::new();
        for path in stream_names {
            if !path.starts_with("BodyText/Section") {
                continue;
            }
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
                diagnostics,
            ) {
                Some(bytes) => bytes,
                None => continue,
            };
            let data = match maybe_decompress_stream(&data, header_ctx.compressed, path) {
                Ok(bytes) => bytes,
                Err(err) => {
                    push_warning(
                        diagnostics,
                        "HWP_DECOMPRESS_FAIL",
                        err.to_string(),
                        Some(path),
                    );
                    continue;
                }
            };
            let paragraph_ids = parse_hwp_section_stream(&data, path, store)?;
            let mut section = Section::new();
            section.name = Some(path.clone());
            section.content = paragraph_ids;
            section.span = Some(SourceSpan::new(path));
            let section_id = section.id;
            store.insert(IRNode::Section(section));
            sections.push(section_id);
        }
        Ok(sections)
    }

    pub(super) fn parse_default_script(&self, cfb: &Cfb, store: &mut IrStore) {
        if let Some(script_data) = cfb.read_stream("Scripts/DefaultJScript") {
            let _ = parse_default_jscript(&script_data, store, "Scripts/DefaultJScript");
        }
    }

    pub(super) fn scan_external_references(
        &self,
        cfb: &Cfb,
        stream_names: &[String],
        header_ctx: &HwpHeaderContext<'_>,
        store: &mut IrStore,
        diagnostics: &mut Diagnostics,
    ) {
        if !header_ctx.allow_parse {
            return;
        }

        let externals = scan_hwp_external_refs(cfb, stream_names, header_ctx, diagnostics);
        for ext in externals {
            store.insert(IRNode::ExternalReference(ext));
        }
    }

    pub(super) fn collect_ole_objects(
        &self,
        cfb: &Cfb,
        stream_names: &[String],
        store: &mut IrStore,
    ) {
        for path in stream_names {
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
    }

    pub(super) fn build_header_context<'a>(
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
