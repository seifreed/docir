use docir_parser::ole::{CfbDirectoryState, CfbEntryType};

pub(super) fn map_state(state: CfbDirectoryState) -> &'static str {
    match state {
        CfbDirectoryState::Normal => "normal",
        CfbDirectoryState::Free => "free",
        CfbDirectoryState::Orphaned => "orphaned",
    }
}

pub(super) fn map_entry_type(entry_type: CfbEntryType) -> &'static str {
    match entry_type {
        CfbEntryType::RootStorage => "root-storage",
        CfbEntryType::Storage => "storage",
        CfbEntryType::Stream => "stream",
    }
}

pub(super) fn classify_entry(path: &str, entry_type: CfbEntryType) -> String {
    let upper = path.to_ascii_uppercase();
    match entry_type {
        CfbEntryType::RootStorage => "root-storage".to_string(),
        CfbEntryType::Storage => {
            if upper == "OBJECTPOOL" || upper.starts_with("OBJECTPOOL/") {
                "embedded-object-storage".to_string()
            } else if upper == "VBA" || upper.ends_with("/VBA") {
                "vba-storage".to_string()
            } else {
                "storage".to_string()
            }
        }
        CfbEntryType::Stream => {
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
            } else if upper == "PACKAGE" || upper.ends_with("/PACKAGE") {
                "package-payload".to_string()
            } else if upper.ends_with("/CONTENTS") {
                "embedded-contents".to_string()
            } else {
                "stream".to_string()
            }
        }
    }
}