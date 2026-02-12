//! ODF manifest parsing helpers.

use crate::error::ParseError;
use base64::engine::general_purpose::STANDARD;
use base64::Engine;
use quick_xml::events::Event;
use quick_xml::Reader;

#[derive(Debug, Clone)]
pub struct OdfManifestEntry {
    pub path: String,
    pub media_type: Option<String>,
    pub encryption: Option<OdfEncryptionData>,
}

#[derive(Debug, Clone)]
pub struct OdfEncryptionData {
    pub checksum_type: Option<String>,
    pub checksum: Option<Vec<u8>>,
    pub algorithm_name: Option<String>,
    pub init_vector: Option<Vec<u8>>,
    pub key_derivation_name: Option<String>,
    pub salt: Option<Vec<u8>>,
    pub iteration_count: Option<u32>,
    pub key_size: Option<u32>,
}

impl Default for OdfEncryptionData {
    fn default() -> Self {
        Self {
            checksum_type: None,
            checksum: None,
            algorithm_name: None,
            init_vector: None,
            key_derivation_name: None,
            salt: None,
            iteration_count: None,
            key_size: None,
        }
    }
}

pub fn parse_manifest(xml: &str) -> Result<Vec<OdfManifestEntry>, ParseError> {
    let mut entries = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut current_entry: Option<OdfManifestEntry> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"manifest:file-entry" => {
                    current_entry = Some(parse_manifest_entry(&e));
                }
                b"manifest:encryption-data" => {
                    if let Some(entry) = current_entry.as_mut() {
                        let mut enc = entry.encryption.take().unwrap_or_default();
                        apply_encryption_data_attrs(&mut enc, &e);
                        entry.encryption = Some(enc);
                    }
                }
                b"manifest:algorithm" => {
                    if let Some(entry) = current_entry.as_mut() {
                        let mut enc = entry.encryption.take().unwrap_or_default();
                        apply_algorithm_attrs(&mut enc, &e);
                        entry.encryption = Some(enc);
                    }
                }
                b"manifest:key-derivation" => {
                    if let Some(entry) = current_entry.as_mut() {
                        let mut enc = entry.encryption.take().unwrap_or_default();
                        apply_key_derivation_attrs(&mut enc, &e);
                        entry.encryption = Some(enc);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"manifest:file-entry" => {
                    entries.push(parse_manifest_entry(&e));
                }
                b"manifest:encryption-data" => {
                    if let Some(entry) = current_entry.as_mut() {
                        let mut enc = entry.encryption.take().unwrap_or_default();
                        apply_encryption_data_attrs(&mut enc, &e);
                        entry.encryption = Some(enc);
                    }
                }
                b"manifest:algorithm" => {
                    if let Some(entry) = current_entry.as_mut() {
                        let mut enc = entry.encryption.take().unwrap_or_default();
                        apply_algorithm_attrs(&mut enc, &e);
                        entry.encryption = Some(enc);
                    }
                }
                b"manifest:key-derivation" => {
                    if let Some(entry) = current_entry.as_mut() {
                        let mut enc = entry.encryption.take().unwrap_or_default();
                        apply_key_derivation_attrs(&mut enc, &e);
                        entry.encryption = Some(enc);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"manifest:file-entry" {
                    if let Some(entry) = current_entry.take() {
                        entries.push(entry);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "META-INF/manifest.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(entries)
}

fn parse_manifest_entry(e: &quick_xml::events::BytesStart<'_>) -> OdfManifestEntry {
    let path = super::attr_value(e, b"manifest:full-path").unwrap_or_default();
    let media_type = super::attr_value(e, b"manifest:media-type");
    OdfManifestEntry {
        path,
        media_type,
        encryption: None,
    }
}

fn apply_encryption_data_attrs(enc: &mut OdfEncryptionData, e: &quick_xml::events::BytesStart<'_>) {
    enc.checksum_type = super::attr_value(e, b"manifest:checksum-type");
    enc.checksum = super::attr_value(e, b"manifest:checksum").and_then(|v| decode_base64_bytes(&v));
}

fn apply_algorithm_attrs(enc: &mut OdfEncryptionData, e: &quick_xml::events::BytesStart<'_>) {
    enc.algorithm_name = super::attr_value(e, b"manifest:algorithm-name");
    enc.init_vector = super::attr_value(e, b"manifest:initialisation-vector")
        .and_then(|v| decode_base64_bytes(&v));
    enc.key_size = super::attr_value(e, b"manifest:key-size").and_then(|v| v.parse::<u32>().ok());
}

fn apply_key_derivation_attrs(enc: &mut OdfEncryptionData, e: &quick_xml::events::BytesStart<'_>) {
    enc.key_derivation_name = super::attr_value(e, b"manifest:key-derivation-name");
    enc.salt = super::attr_value(e, b"manifest:salt").and_then(|v| decode_base64_bytes(&v));
    enc.iteration_count =
        super::attr_value(e, b"manifest:iteration-count").and_then(|v| v.parse::<u32>().ok());
}

pub fn is_manifest_entry_encrypted(entry: &OdfManifestEntry) -> bool {
    if entry.encryption.is_some() {
        return true;
    }
    if let Some(media) = entry.media_type.as_deref() {
        if media.contains("encrypted") {
            return true;
        }
    }
    entry.path.to_ascii_lowercase().contains("encrypted")
}

pub fn manifest_has_encryption(entries: &[OdfManifestEntry]) -> bool {
    entries.iter().any(is_manifest_entry_encrypted)
}

pub fn encrypted_manifest_entries(entries: &[OdfManifestEntry]) -> Vec<String> {
    entries
        .iter()
        .filter(|entry| is_manifest_entry_encrypted(entry))
        .map(|entry| entry.path.clone())
        .collect()
}

pub fn format_odf_encryption_metadata(entry: &OdfManifestEntry) -> Option<String> {
    let enc = entry.encryption.as_ref()?;
    let algorithm = enc.algorithm_name.as_deref().unwrap_or("unknown");
    let kdf = enc.key_derivation_name.as_deref().unwrap_or("unknown");
    let key_bits = enc
        .key_size
        .map(|v| v.to_string())
        .unwrap_or_else(|| "unknown".to_string());
    let iterations = enc
        .iteration_count
        .map(|v| v.to_string())
        .unwrap_or_else(|| "unknown".to_string());
    let iv = enc
        .init_vector
        .as_ref()
        .map(|v| STANDARD.encode(v))
        .unwrap_or_else(|| "unknown".to_string());
    let checksum = enc
        .checksum
        .as_ref()
        .map(|v| STANDARD.encode(v))
        .unwrap_or_else(|| "unknown".to_string());
    let checksum_type = enc.checksum_type.as_deref().unwrap_or("unknown");
    Some(format!(
        "ODF encryption: algorithm={algorithm}, kdf={kdf}, key_bits={key_bits}, iterations={iterations}, iv={iv}, checksum={checksum} ({checksum_type})"
    ))
}

fn decode_base64_bytes(value: &str) -> Option<Vec<u8>> {
    STANDARD.decode(value.as_bytes()).ok()
}
