use crate::{AppError, ParsedDocument};
use docir_core::types::DocumentFormat;
use docir_parser::ole::{Cfb, CfbEntryType};
use docir_parser::zip_handler::{SecureZipReader, ZipConfig};
use serde::Serialize;
use std::io::Cursor;

use crate::inventory::ContainerKind;

/// Serializable dump of low-level container entries for analyst-facing inspection.
#[derive(Debug, Clone, Serialize)]
pub struct ContainerDump {
    pub document_format: String,
    pub container_kind: ContainerKind,
    pub entry_count: usize,
    pub entries: Vec<ContainerEntry>,
}

/// One physical entry surfaced from the source container.
#[derive(Debug, Clone, Serialize)]
pub struct ContainerEntry {
    pub kind: ContainerEntryKind,
    pub path: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start_sector: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub created_filetime: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub modified_filetime: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub details: Option<String>,
}

/// Low-level entry taxonomy for raw container dumps.
#[derive(Debug, Clone, Copy, Serialize, PartialEq, Eq)]
pub enum ContainerEntryKind {
    ZipEntry,
    CfbRootStorage,
    CfbStorage,
    CfbStream,
    RtfDocument,
    RtfEmbeddedObject,
}

impl ContainerDump {
    /// Builds a container dump from original source bytes and parsed format context.
    pub fn from_parsed_bytes(
        parsed: &ParsedDocument,
        input_bytes: &[u8],
        zip_config: &ZipConfig,
    ) -> Result<Self, AppError> {
        let mut dump = Self {
            document_format: parsed.format().extension().to_string(),
            container_kind: classify_container(parsed),
            entry_count: 0,
            entries: Vec::new(),
        };

        match dump.container_kind {
            ContainerKind::ZipOoxml | ContainerKind::ZipOdf | ContainerKind::ZipHwpx => {
                let mut zip = SecureZipReader::new(Cursor::new(input_bytes), zip_config.clone())?;
                let mut names: Vec<String> = zip.file_names().map(str::to_string).collect();
                names.sort();
                for path in names {
                    dump.entries.push(ContainerEntry {
                        kind: ContainerEntryKind::ZipEntry,
                        size_bytes: zip.file_size(&path).ok(),
                        start_sector: None,
                        created_filetime: None,
                        modified_filetime: None,
                        path,
                        details: None,
                    });
                }
            }
            ContainerKind::CfbOle => {
                let cfb = Cfb::parse(input_bytes.to_vec())?;
                for entry in cfb.list_entries() {
                    dump.entries.push(ContainerEntry {
                        kind: map_cfb_entry_kind(entry.entry_type),
                        path: entry.path.clone(),
                        size_bytes: Some(entry.size),
                        start_sector: Some(entry.start_sector),
                        created_filetime: entry.created_filetime,
                        modified_filetime: entry.modified_filetime,
                        details: classify_cfb_entry(&entry.path, entry.entry_type),
                    });
                }
            }
            ContainerKind::Rtf => {
                dump.entries.push(ContainerEntry {
                    kind: ContainerEntryKind::RtfDocument,
                    path: "rtf:/document".to_string(),
                    size_bytes: Some(input_bytes.len() as u64),
                    start_sector: None,
                    created_filetime: None,
                    modified_filetime: None,
                    details: None,
                });
                for (idx, blob) in scan_rtf_objdata(input_bytes).into_iter().enumerate() {
                    dump.entries.push(ContainerEntry {
                        kind: ContainerEntryKind::RtfEmbeddedObject,
                        path: format!("rtf:/objdata/{}", idx + 1),
                        size_bytes: Some(blob.len() as u64),
                        start_sector: None,
                        created_filetime: None,
                        modified_filetime: None,
                        details: Some(classify_rtf_blob(&blob).to_string()),
                    });
                }
            }
            ContainerKind::Unknown => {
                dump.entries.push(ContainerEntry {
                    kind: ContainerEntryKind::ZipEntry,
                    path: "unknown:/".to_string(),
                    size_bytes: Some(input_bytes.len() as u64),
                    start_sector: None,
                    created_filetime: None,
                    modified_filetime: None,
                    details: Some("opaque container".to_string()),
                });
            }
        }

        dump.entry_count = dump.entries.len();
        Ok(dump)
    }
}

fn classify_container(parsed: &ParsedDocument) -> ContainerKind {
    if parsed
        .document()
        .and_then(|doc| doc.span.as_ref())
        .map(|span| span.file_path.starts_with("cfb:/"))
        .unwrap_or(false)
    {
        return ContainerKind::CfbOle;
    }

    match parsed.format() {
        DocumentFormat::WordProcessing
        | DocumentFormat::Spreadsheet
        | DocumentFormat::Presentation => ContainerKind::ZipOoxml,
        DocumentFormat::OdfText
        | DocumentFormat::OdfSpreadsheet
        | DocumentFormat::OdfPresentation => ContainerKind::ZipOdf,
        DocumentFormat::Hwpx => ContainerKind::ZipHwpx,
        DocumentFormat::Hwp => ContainerKind::CfbOle,
        DocumentFormat::Rtf => ContainerKind::Rtf,
    }
}

fn map_cfb_entry_kind(entry_type: CfbEntryType) -> ContainerEntryKind {
    match entry_type {
        CfbEntryType::RootStorage => ContainerEntryKind::CfbRootStorage,
        CfbEntryType::Storage => ContainerEntryKind::CfbStorage,
        CfbEntryType::Stream => ContainerEntryKind::CfbStream,
    }
}

fn classify_cfb_entry(path: &str, entry_type: CfbEntryType) -> Option<String> {
    let upper = path.to_ascii_uppercase();
    if matches!(entry_type, CfbEntryType::RootStorage) {
        return Some("root-storage".to_string());
    }
    if upper == "WORDDOCUMENT" {
        return Some("word-main-stream".to_string());
    }
    if upper == "WORKBOOK" || upper == "BOOK" {
        return Some("excel-main-stream".to_string());
    }
    if upper == "POWERPOINT DOCUMENT" {
        return Some("powerpoint-main-stream".to_string());
    }
    if upper.ends_with("/PROJECT") || upper == "PROJECT" {
        return Some("vba-project-metadata".to_string());
    }
    if upper.contains("/VBA/") || upper.starts_with("VBA/") {
        return Some("vba-module-stream".to_string());
    }
    if upper.ends_with("OLE10NATIVE") {
        return Some("ole-native-payload".to_string());
    }
    if upper.ends_with("/PACKAGE") {
        return Some("package-payload".to_string());
    }
    if upper == "OBJECTPOOL" || upper.starts_with("OBJECTPOOL/") {
        return Some("embedded-object-storage".to_string());
    }
    if upper.ends_with("/CONTENTS") {
        return Some("embedded-contents".to_string());
    }
    None
}

fn classify_rtf_blob(blob: &[u8]) -> &'static str {
    if blob.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]) {
        return "ole-object";
    }
    if blob.starts_with(b"MZ") {
        return "pe";
    }
    if blob.starts_with(b"%PDF-") {
        return "pdf";
    }
    "blob"
}

fn scan_rtf_objdata(data: &[u8]) -> Vec<Vec<u8>> {
    let mut blobs = Vec::new();
    let mut index = 0usize;
    let mut depth = 0usize;
    let mut capture_depth = None::<usize>;
    let mut hex = Vec::new();

    while index < data.len() {
        match data[index] {
            b'{' => {
                depth = depth.saturating_add(1);
                index += 1;
            }
            b'}' => {
                if let Some(target_depth) = capture_depth {
                    if depth <= target_depth {
                        if let Some(blob) = decode_hex_blob(&hex) {
                            blobs.push(blob);
                        }
                        hex.clear();
                        capture_depth = None;
                    }
                }
                depth = depth.saturating_sub(1);
                index += 1;
            }
            b'\\' => {
                index += 1;
                let start = index;
                while index < data.len() && data[index].is_ascii_alphabetic() {
                    index += 1;
                }
                let word = std::str::from_utf8(&data[start..index]).unwrap_or("");
                if word == "objdata" {
                    capture_depth = Some(depth);
                    hex.clear();
                }
                if index < data.len() && (data[index] == b'-' || data[index].is_ascii_digit()) {
                    index += 1;
                    while index < data.len() && data[index].is_ascii_digit() {
                        index += 1;
                    }
                }
                if index < data.len() && data[index] == b' ' {
                    index += 1;
                }
            }
            byte => {
                if capture_depth.is_some() && byte.is_ascii_hexdigit() {
                    hex.push(byte);
                }
                index += 1;
            }
        }
    }

    if capture_depth.is_some() {
        if let Some(blob) = decode_hex_blob(&hex) {
            blobs.push(blob);
        }
    }

    blobs
}

fn decode_hex_blob(hex: &[u8]) -> Option<Vec<u8>> {
    if hex.len() < 2 {
        return None;
    }
    let even_len = hex.len() - (hex.len() % 2);
    let mut out = Vec::with_capacity(even_len / 2);
    let mut index = 0usize;
    while index + 1 < even_len {
        let hi = hex_val(hex[index])?;
        let lo = hex_val(hex[index + 1])?;
        out.push((hi << 4) | lo);
        index += 2;
    }
    Some(out)
}

fn hex_val(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{test_support::build_test_cfb, ParseMetrics, ParsedDocument};
    use docir_core::ir::{Document, IRNode};
    use docir_core::types::{DocumentFormat, SourceSpan};
    use docir_core::visitor::IrStore;

    fn parsed_document(format: DocumentFormat, span: Option<&str>) -> ParsedDocument {
        let mut store = IrStore::new();
        let mut document = Document::new(format);
        document.span = span.map(SourceSpan::new);
        let root_id = document.id;
        store.insert(IRNode::Document(document));
        ParsedDocument::new(docir_parser::parser::ParsedDocument {
            root_id,
            format,
            store,
            metrics: Some(ParseMetrics::default()),
        })
    }

    #[test]
    fn container_dump_lists_zip_entries() {
        let parsed = parsed_document(DocumentFormat::WordProcessing, None);
        let bytes = std::fs::read(
            std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                .join("../../fixtures/ooxml/minimal.docx"),
        )
        .expect("fixture");
        let dump =
            ContainerDump::from_parsed_bytes(&parsed, &bytes, &ZipConfig::default()).expect("zip");
        assert_eq!(dump.container_kind, ContainerKind::ZipOoxml);
        assert!(dump
            .entries
            .iter()
            .any(|entry| entry.path == "word/document.xml"));
    }

    #[test]
    fn container_dump_lists_cfb_streams() {
        let parsed = parsed_document(DocumentFormat::WordProcessing, Some("cfb:/"));
        let bytes = build_test_cfb_fixture(&[
            ("WordDocument", b"main"),
            ("Macros/PROJECT", b"ID=\"VBAProject\""),
            ("ObjectPool/1/Ole10Native", b"payload"),
        ]);
        let dump =
            ContainerDump::from_parsed_bytes(&parsed, &bytes, &ZipConfig::default()).expect("cfb");
        assert_eq!(dump.container_kind, ContainerKind::CfbOle);
        assert!(dump.entries.iter().any(|entry| {
            entry.path.ends_with("Ole10Native")
                && entry.details.as_deref() == Some("ole-native-payload")
        }));
    }

    #[test]
    fn container_dump_lists_rtf_objects() {
        let parsed = parsed_document(DocumentFormat::Rtf, None);
        let bytes = br"{\rtf1{\object{\objdata 4d5a9000}}}".to_vec();
        let dump =
            ContainerDump::from_parsed_bytes(&parsed, &bytes, &ZipConfig::default()).expect("rtf");
        assert_eq!(dump.container_kind, ContainerKind::Rtf);
        assert!(dump
            .entries
            .iter()
            .any(|entry| entry.path == "rtf:/objdata/1"));
    }

    fn build_test_cfb_fixture(entries: &[(&str, &[u8])]) -> Vec<u8> {
        build_test_cfb(entries)
    }
}
