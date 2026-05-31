use super::{InventoryArtifact, InventoryArtifactKind};
use docir_parser::ole::{Cfb, CfbEntryType};

pub(super) fn build_cfb_container_entries(input_bytes: &[u8]) -> Vec<InventoryArtifact> {
    let Ok(cfb) = Cfb::parse(input_bytes.to_vec()) else {
        return Vec::new();
    };

    cfb.list_entries()
        .into_iter()
        .map(|entry| InventoryArtifact {
            kind: InventoryArtifactKind::ContainerEntry,
            node_id: None,
            path: Some(entry.path.clone()),
            relationship_id: None,
            size_bytes: Some(entry.size),
            start_sector: Some(entry.start_sector),
            created_filetime: entry.created_filetime,
            modified_filetime: entry.modified_filetime,
            media_type: None,
            sha256: None,
            details: classify_cfb_entry_detail(&entry.path, entry.entry_type),
        })
        .collect()
}

fn classify_cfb_entry_detail(path: &str, entry_type: CfbEntryType) -> String {
    match entry_type {
        CfbEntryType::RootStorage => "root-storage".to_string(),
        CfbEntryType::Storage => classify_cfb_storage(path),
        CfbEntryType::Stream => classify_cfb_stream(path),
    }
}

fn classify_cfb_storage(path: &str) -> String {
    let upper = path.to_ascii_uppercase();
    if upper == "OBJECTPOOL" || upper.starts_with("OBJECTPOOL/") {
        "embedded-object-storage".to_string()
    } else if upper == "VBA" || upper.ends_with("/VBA") {
        "vba-storage".to_string()
    } else {
        "storage".to_string()
    }
}

fn classify_cfb_stream(path: &str) -> String {
    let upper = path.to_ascii_uppercase();
    if upper == "WORDDOCUMENT" {
        "word-main-stream".to_string()
    } else if upper == "WORKBOOK" || upper == "BOOK" {
        "excel-main-stream".to_string()
    } else if upper == "POWERPOINT DOCUMENT" {
        "powerpoint-main-stream".to_string()
    } else if upper.ends_with("/PROJECT") || upper == "PROJECT" {
        "vba-project-metadata".to_string()
    } else if upper.contains("/VBA/") || upper.starts_with("VBA/") {
        "vba-module-stream".to_string()
    } else if upper.ends_with("OLE10NATIVE") {
        "ole-native-payload".to_string()
    } else if upper.ends_with("/PACKAGE") {
        "package-payload".to_string()
    } else if upper.ends_with("/CONTENTS") {
        "embedded-contents".to_string()
    } else {
        "stream".to_string()
    }
}
