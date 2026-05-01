use super::super::helpers::{
    is_hwpx_footer, is_hwpx_header, is_hwpx_master, is_hwpx_section, media_type_from_path,
};
use super::super::section::parse_hwpx_section;
use super::super::styles::parse_hwpx_styles;
use super::super::HwpxParser;
use crate::error::ParseError;
use crate::zip_handler::SecureZipReader;
use docir_core::ir::{Document, Footer, Header, IRNode, Section};
use docir_core::ir::{ExtensionPart, ExtensionPartKind, MediaAsset};
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use std::collections::HashMap;
use std::io::{Read, Seek};

impl HwpxParser {
    pub(super) fn collect_hwpx_parts<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        file_names: &[String],
        store: &mut IrStore,
    ) -> HwpxPartCollection {
        let mut shared_parts = Vec::new();
        let mut media_assets = Vec::new();
        let mut media_lookup: HashMap<String, NodeId> = HashMap::new();
        for path in file_names {
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
        HwpxPartCollection {
            shared_parts,
            media_assets,
            media_lookup,
        }
    }

    pub(super) fn parse_hwpx_primary_sections<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        file_names: &[String],
        media_lookup: &HashMap<String, NodeId>,
        store: &mut IrStore,
        doc: &mut Document,
    ) -> Result<Vec<NodeId>, ParseError> {
        let mut section_ids = Vec::new();
        for path in file_names {
            if !is_hwpx_section(path) {
                continue;
            }
            let content = self.parse_hwpx_section_content(zip, path, media_lookup, store, doc)?;
            let mut section = Section::new();
            section.name = Some(path.clone());
            section.content = content;
            section.span = Some(SourceSpan::new(path));
            let section_id = section.id;
            store.insert(IRNode::Section(section));
            section_ids.push(section_id);
        }
        Ok(section_ids)
    }

    pub(super) fn parse_hwpx_section_groups<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        file_names: &[String],
        media_lookup: &HashMap<String, NodeId>,
        store: &mut IrStore,
        doc: &mut Document,
    ) -> Result<HwpxSectionGroups, ParseError> {
        let mut grouped = HwpxSectionGroups::default();
        for path in file_names {
            if is_hwpx_header(path) {
                let content =
                    self.parse_hwpx_section_content(zip, path, media_lookup, store, doc)?;
                let mut header = Header::new();
                header.content = content;
                header.span = Some(SourceSpan::new(path));
                let header_id = header.id;
                store.insert(IRNode::Header(header));
                grouped.header_ids.push(header_id);
            } else if is_hwpx_footer(path) {
                let content =
                    self.parse_hwpx_section_content(zip, path, media_lookup, store, doc)?;
                let mut footer = Footer::new();
                footer.content = content;
                footer.span = Some(SourceSpan::new(path));
                let footer_id = footer.id;
                store.insert(IRNode::Footer(footer));
                grouped.footer_ids.push(footer_id);
            } else if is_hwpx_master(path) {
                let content =
                    self.parse_hwpx_section_content(zip, path, media_lookup, store, doc)?;
                let mut section = Section::new();
                section.name = Some(format!("master:{}", path));
                section.content = content;
                section.span = Some(SourceSpan::new(path));
                let section_id = section.id;
                store.insert(IRNode::Section(section));
                grouped.master_ids.push(section_id);
            }
        }
        Ok(grouped)
    }

    pub(super) fn parse_hwpx_section_content<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        path: &str,
        media_lookup: &HashMap<String, NodeId>,
        store: &mut IrStore,
        doc: &mut Document,
    ) -> Result<Vec<NodeId>, ParseError> {
        let xml = zip.read_file_string(path)?;
        parse_hwpx_section(
            &xml,
            path,
            store,
            &mut doc.comments,
            &mut doc.footnotes,
            &mut doc.endnotes,
            media_lookup,
        )
    }

    pub(super) fn attach_hwpx_headers_and_footers(
        &self,
        store: &mut IrStore,
        section_ids: &[NodeId],
        header_ids: &[NodeId],
        footer_ids: &[NodeId],
    ) {
        for section_id in section_ids {
            if let Some(IRNode::Section(section)) = store.get_mut(*section_id) {
                section.headers.extend(header_ids.iter().copied());
                section.footers.extend(footer_ids.iter().copied());
            }
        }
    }

    pub(super) fn parse_hwpx_styles_part<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
        doc: &mut Document,
    ) {
        if !zip.contains("Contents/content.hpf") {
            return;
        }
        if let Ok(xml) = zip.read_file_string("Contents/content.hpf") {
            if let Some(style_set) = parse_hwpx_styles(&xml, "Contents/content.hpf") {
                let style_id = style_set.id;
                store.insert(IRNode::StyleSet(style_set));
                doc.styles = Some(style_id);
            }
        }
    }
}

pub(super) struct HwpxPartCollection {
    pub(super) shared_parts: Vec<NodeId>,
    pub(super) media_assets: Vec<NodeId>,
    pub(super) media_lookup: HashMap<String, NodeId>,
}

#[derive(Default)]
pub(super) struct HwpxSectionGroups {
    pub(super) header_ids: Vec<NodeId>,
    pub(super) footer_ids: Vec<NodeId>,
    pub(super) master_ids: Vec<NodeId>,
}
