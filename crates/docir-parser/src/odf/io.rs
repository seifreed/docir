use super::{build_media_asset, OdfManifestEntry};
use crate::diagnostics::push_info;
use crate::zip_handler::SecureZipReader;
use docir_core::ir::{Document, ExtensionPart, ExtensionPartKind, IRNode};
use docir_core::visitor::IrStore;
use std::collections::HashMap;
use std::io::{Read, Seek};

pub(super) fn collect_manifest_index(
    manifest_entries: &[OdfManifestEntry],
    diagnostics: &mut docir_core::ir::Diagnostics,
) -> HashMap<String, Option<String>> {
    let mut manifest_index: HashMap<String, Option<String>> = HashMap::new();
    for entry in manifest_entries {
        let path = entry.path.clone();
        let media_type = entry.media_type.clone();
        manifest_index.insert(path.clone(), media_type.clone());
        push_info(
            diagnostics,
            "ODF_PART",
            format!(
                "ODF part: {} (media-type: {})",
                path,
                media_type.clone().unwrap_or_else(|| "(none)".to_string())
            ),
            Some(&path),
        );
    }
    manifest_index
}

pub(super) fn collect_shared_parts<R: Read + Seek>(
    zip: &mut SecureZipReader<R>,
    manifest_index: &HashMap<String, Option<String>>,
    store: &mut IrStore,
    doc: &mut Document,
) -> Vec<String> {
    let mut file_names: Vec<String> = zip.file_names().map(|name| name.to_string()).collect();
    file_names.sort();
    for path in &file_names {
        if path == "mimetype" {
            continue;
        }
        let media_type = manifest_index.get(path.as_str()).cloned().unwrap_or(None);
        let size_bytes = zip.file_size(path).unwrap_or(0);
        let mut part = ExtensionPart::new(path.to_string(), size_bytes, ExtensionPartKind::Unknown);
        part.content_type = media_type.clone();
        let part_id = part.id;
        store.insert(IRNode::ExtensionPart(part));
        doc.add_shared_part(part_id);

        if let Some(media) = media_type.as_deref() {
            if let Some(asset) = build_media_asset(path, media, size_bytes) {
                let asset_id = asset.id;
                store.insert(IRNode::MediaAsset(asset));
                doc.add_shared_part(asset_id);
            }
        }
    }
    file_names
}
