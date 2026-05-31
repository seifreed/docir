use super::vba_scanner::VbaScanner;
use super::{ParseError, ParserConfig, hex};
use crate::ole::Cfb;
use crate::ooxml::part_utils::get_rels_path;
use crate::ooxml::relationships::{Relationships, TargetMode, rel_type};
use crate::zip_handler::PackageReader;
use docir_core::ir::IRNode;
use docir_core::security::{ExternalRefType, ExternalReference, OleObject};
use docir_core::types::SourceSpan;
use docir_core::visitor::IrStore;
use std::collections::HashSet;

pub struct SecurityScanner<'a> {
    config: &'a ParserConfig,
}

impl<'a> SecurityScanner<'a> {
    /// Public API entrypoint: new.
    pub fn new(config: &'a ParserConfig) -> Self {
        Self { config }
    }

    pub fn scan_zip(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        VbaScanner::new(self.config).scan_zip_vba_projects(zip, store)?;
        self.scan_ole_objects(zip, store)?;
        self.scan_activex_controls(zip, store)?;
        self.scan_word_external_relationships(zip, store)?;
        Ok(())
    }

    pub fn scan_cfb(&self, cfb: &Cfb, store: &mut IrStore) -> Result<(), ParseError> {
        let streams = cfb.list_streams();
        VbaScanner::new(self.config).scan_cfb_vba_projects(cfb, &streams, store)?;
        self.scan_cfb_ole_objects(cfb, &streams, store)?;
        Ok(())
    }

    fn scan_ole_objects(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        let mut builder = docir_core::ir::IrBuilder::new(store);
        let ole_files: Vec<String> = zip
            .list_prefix("word/embeddings/")
            .into_iter()
            .chain(zip.list_prefix("xl/embeddings/"))
            .chain(zip.list_prefix("ppt/embeddings/"))
            .filter(|p| p.ends_with(".bin") || p.ends_with(".ole"))
            .collect();

        for ole_path in ole_files {
            let ole_object = self.detect_ole_object(zip, &ole_path)?;
            builder.insert(IRNode::OleObject(ole_object));
        }
        Ok(())
    }

    fn scan_activex_controls(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        let mut activex_bin_seen: HashSet<String> = HashSet::new();
        let activex_paths: Vec<String> = zip
            .list_prefix("word/activeX/")
            .into_iter()
            .chain(zip.list_prefix("xl/activeX/"))
            .chain(zip.list_prefix("ppt/activeX/"))
            .filter(|p| p.ends_with(".xml"))
            .collect();
        for path in activex_paths {
            let xml = zip.read_file_string(&path)?;
            if let Some(mut control) = super::parse_activex_xml(&xml, &path) {
                control.span = Some(SourceSpan::new(&path));
                store.insert(IRNode::ActiveXControl(control));
            }
            self.scan_activex_control_rels(zip, store, &path, &mut activex_bin_seen)?;
        }
        Ok(())
    }

    fn scan_word_external_relationships(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        let rel_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.starts_with("word/") && p.ends_with(".rels"))
            .collect();
        for rel_path in rel_paths {
            let rels_xml = zip.read_file_string(&rel_path)?;
            let rels = Relationships::parse(&rels_xml)?;
            for rel in rels.by_id.values() {
                if rel.target_mode == TargetMode::External {
                    self.insert_external_ref(store, &rel_path, rel);
                }
            }
        }
        Ok(())
    }

    fn scan_activex_control_rels(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
        control_path: &str,
        activex_bin_seen: &mut HashSet<String>,
    ) -> Result<(), ParseError> {
        let rels_path = get_rels_path(control_path);
        if !zip.contains(&rels_path) {
            return Ok(());
        }

        let Ok(rels_xml) = zip.read_file_string(&rels_path) else {
            return Ok(());
        };
        let Ok(rels) = Relationships::parse(&rels_xml) else {
            return Ok(());
        };

        for rel in rels.by_id.values() {
            if !self.is_activex_binary_rel(rel) {
                continue;
            }
            let bin_path = Relationships::resolve_target(control_path, &rel.target);
            if activex_bin_seen.insert(bin_path.clone()) && zip.contains(&bin_path) {
                let ole_object = self.detect_ole_object(zip, &bin_path)?;
                store.insert(IRNode::OleObject(ole_object));
            }
        }

        Ok(())
    }

    fn is_activex_binary_rel(&self, rel: &crate::ooxml::relationships::Relationship) -> bool {
        rel.target.ends_with(".bin") || rel.rel_type.contains("activeXControlBinary")
    }

    fn insert_external_ref(
        &self,
        store: &mut IrStore,
        rel_path: &str,
        rel: &crate::ooxml::relationships::Relationship,
    ) {
        let ref_type = map_external_ref_type(&rel.rel_type);
        let mut ext_ref = ExternalReference::new(ref_type, &rel.target);
        ext_ref.relationship_id = Some(rel.id.clone());
        ext_ref.relationship_type = Some(rel.rel_type.clone());
        ext_ref.span = Some(SourceSpan::new(rel_path));
        store.insert(IRNode::ExternalReference(ext_ref));
    }

    fn detect_ole_object(
        &self,
        zip: &mut impl PackageReader,
        path: &str,
    ) -> Result<OleObject, ParseError> {
        let data = zip.read_file(path)?;
        Ok(self.build_ole_object_from_bytes(path, &data))
    }

    fn scan_cfb_ole_objects(
        &self,
        cfb: &Cfb,
        streams: &[String],
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        for path in streams.iter().filter(|path| {
            let upper = path.to_ascii_uppercase();
            upper.contains("OBJECTPOOL/")
                || upper.ends_with("OLE10NATIVE")
                || upper.ends_with("/PACKAGE")
                || upper.ends_with("/CONTENTS")
        }) {
            if let Some(bytes) = cfb.read_stream(path) {
                store.insert(IRNode::OleObject(
                    self.build_ole_object_from_bytes(path, &bytes),
                ));
            }
        }
        Ok(())
    }

    fn build_ole_object_from_bytes(&self, path: &str, data: &[u8]) -> OleObject {
        let mut ole = OleObject::new();
        ole.source_path = Some(path.to_string());
        ole.span = Some(SourceSpan::new(path));
        ole.size_bytes = data.len() as u64;
        let upper = path.to_ascii_uppercase();
        if upper.contains("OBJECTPOOL/") {
            ole.class_name = Some("ObjectPool".to_string());
        }
        if upper.ends_with("OLE10NATIVE") {
            ole.embedded_payload_kind = Some("ole10native".to_string());
        } else if upper.ends_with("/PACKAGE") || upper == "PACKAGE" {
            ole.embedded_payload_kind = Some("package".to_string());
        }

        if self.config.compute_hashes {
            ole.data_hash = Some(compute_sha256_hex(data));
        }

        ole
    }
}

fn compute_sha256_hex(data: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    let mut hasher = Sha256::new();
    hasher.update(data);
    hex::encode(hasher.finalize())
}

fn map_external_ref_type(rel_type_value: &str) -> ExternalRefType {
    match rel_type_value {
        rel_type::HYPERLINK => ExternalRefType::Hyperlink,
        rel_type::IMAGE => ExternalRefType::Image,
        rel_type::OLE_OBJECT => ExternalRefType::OleLink,
        rel_type::ATTACHED_TEMPLATE => ExternalRefType::AttachedTemplate,
        _ => ExternalRefType::Other,
    }
}

#[cfg(test)]
#[path = "security_tests.rs"]
mod tests;
