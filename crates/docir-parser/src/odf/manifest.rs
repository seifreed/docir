//! ODF manifest parsing helpers.

use crate::error::ParseError;
use crate::xml_utils::{
    XmlScanControl, local_name, scan_xml_events, try_attr_value_by_suffix, xml_error,
};
use base64::Engine;
use base64::engine::general_purpose::STANDARD;
use quick_xml::Reader;
use quick_xml::events::Event;

#[derive(Debug, Clone)]
pub struct OdfManifestEntry {
    pub path: String,
    pub media_type: Option<String>,
    pub encryption: Option<OdfEncryptionData>,
}

#[derive(Debug, Clone, Default)]
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

/// Public API entrypoint: parse_manifest.
pub fn parse_manifest(xml: &str) -> Result<Vec<OdfManifestEntry>, ParseError> {
    let mut entries = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut current_entry: Option<OdfManifestEntry> = None;

    scan_xml_events(&mut reader, &mut buf, "META-INF/manifest.xml", |event| {
        match event {
            Event::Start(e) => handle_manifest_start_event(&e, &mut current_entry)?,
            Event::Empty(e) => handle_manifest_empty_event(&e, &mut entries, &mut current_entry)?,
            Event::End(e) => {
                handle_manifest_end_event(e.name().as_ref(), &mut entries, &mut current_entry)
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok(entries)
}

fn handle_manifest_start_event(
    e: &quick_xml::events::BytesStart<'_>,
    current_entry: &mut Option<OdfManifestEntry>,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"file-entry" => {
            *current_entry = Some(parse_manifest_entry(e)?);
        }
        b"encryption-data" => {
            apply_entry_encryption_attrs(current_entry, e, apply_encryption_data_attrs)?;
        }
        b"algorithm" => {
            apply_entry_encryption_attrs(current_entry, e, apply_algorithm_attrs)?;
        }
        b"key-derivation" => {
            apply_entry_encryption_attrs(current_entry, e, apply_key_derivation_attrs)?;
        }
        _ => {}
    }
    Ok(())
}

fn handle_manifest_empty_event(
    e: &quick_xml::events::BytesStart<'_>,
    entries: &mut Vec<OdfManifestEntry>,
    current_entry: &mut Option<OdfManifestEntry>,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"file-entry" => entries.push(parse_manifest_entry(e)?),
        b"encryption-data" => {
            apply_entry_encryption_attrs(current_entry, e, apply_encryption_data_attrs)?;
        }
        b"algorithm" => apply_entry_encryption_attrs(current_entry, e, apply_algorithm_attrs)?,
        b"key-derivation" => {
            apply_entry_encryption_attrs(current_entry, e, apply_key_derivation_attrs)?;
        }
        _ => {}
    }
    Ok(())
}

fn handle_manifest_end_event(
    name: &[u8],
    entries: &mut Vec<OdfManifestEntry>,
    current_entry: &mut Option<OdfManifestEntry>,
) {
    if local_name(name) == b"file-entry"
        && let Some(entry) = current_entry.take()
    {
        entries.push(entry);
    }
}

fn apply_entry_encryption_attrs(
    current_entry: &mut Option<OdfManifestEntry>,
    e: &quick_xml::events::BytesStart<'_>,
    apply_fn: fn(
        &mut OdfEncryptionData,
        &quick_xml::events::BytesStart<'_>,
    ) -> Result<(), ParseError>,
) -> Result<(), ParseError> {
    if let Some(entry) = current_entry.as_mut() {
        let mut enc = entry.encryption.take().unwrap_or_default();
        apply_fn(&mut enc, e)?;
        entry.encryption = Some(enc);
    }
    Ok(())
}

fn parse_manifest_entry(
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<OdfManifestEntry, ParseError> {
    let path =
        try_attr_value_by_suffix(e, &[b":full-path"], "META-INF/manifest.xml")?.unwrap_or_default();
    let media_type = try_attr_value_by_suffix(e, &[b":media-type"], "META-INF/manifest.xml")?;
    Ok(OdfManifestEntry {
        path,
        media_type,
        encryption: None,
    })
}

fn apply_encryption_data_attrs(
    enc: &mut OdfEncryptionData,
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), ParseError> {
    enc.checksum_type = try_attr_value_by_suffix(e, &[b":checksum-type"], "META-INF/manifest.xml")?;
    enc.checksum = try_attr_value_by_suffix(e, &[b":checksum"], "META-INF/manifest.xml")?
        .and_then(|v| decode_base64_bytes(&v));
    Ok(())
}

fn apply_algorithm_attrs(
    enc: &mut OdfEncryptionData,
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), ParseError> {
    enc.algorithm_name =
        try_attr_value_by_suffix(e, &[b":algorithm-name"], "META-INF/manifest.xml")?;
    enc.init_vector =
        try_attr_value_by_suffix(e, &[b":initialisation-vector"], "META-INF/manifest.xml")?
            .and_then(|v| decode_base64_bytes(&v));
    enc.key_size = parse_optional_u32_attr(e, &[b":key-size"])?;
    Ok(())
}

fn apply_key_derivation_attrs(
    enc: &mut OdfEncryptionData,
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), ParseError> {
    enc.key_derivation_name =
        try_attr_value_by_suffix(e, &[b":key-derivation-name"], "META-INF/manifest.xml")?;
    enc.salt = try_attr_value_by_suffix(e, &[b":salt"], "META-INF/manifest.xml")?
        .and_then(|v| decode_base64_bytes(&v));
    enc.iteration_count = parse_optional_u32_attr(e, &[b":iteration-count"])?;
    Ok(())
}

fn parse_optional_u32_attr(
    e: &quick_xml::events::BytesStart<'_>,
    names: &[&[u8]],
) -> Result<Option<u32>, ParseError> {
    try_attr_value_by_suffix(e, names, "META-INF/manifest.xml")?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|err| xml_error("META-INF/manifest.xml", err))
        })
        .transpose()
}

/// Public API entrypoint: is_manifest_entry_encrypted.
pub fn is_manifest_entry_encrypted(entry: &OdfManifestEntry) -> bool {
    if entry.encryption.is_some() {
        return true;
    }
    if let Some(media) = entry.media_type.as_deref()
        && media.contains("encrypted")
    {
        return true;
    }
    entry.path.to_ascii_lowercase().contains("encrypted")
}

/// Public API entrypoint: encrypted_manifest_entries.
pub fn encrypted_manifest_entries(entries: &[OdfManifestEntry]) -> Vec<String> {
    entries
        .iter()
        .filter(|entry| is_manifest_entry_encrypted(entry))
        .map(|entry| entry.path.clone())
        .collect()
}

/// Public API entrypoint: format_odf_encryption_metadata.
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

#[cfg(test)]
mod tests {
    use super::{format_odf_encryption_metadata, parse_manifest};

    #[test]
    fn parse_manifest_accepts_alternate_namespace_prefixes() {
        let xml = r#"
            <mf:manifest xmlns:mf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
              <mf:file-entry mf:full-path="content.xml" mf:media-type="text/xml">
                <mf:encryption-data mf:checksum-type="SHA1" mf:checksum="YWJjZA==">
                  <mf:algorithm mf:algorithm-name="http://www.w3.org/2001/04/xmlenc#aes256-cbc"
                    mf:initialisation-vector="MTIzNDU2Nzg5MA==" mf:key-size="32"/>
                  <mf:key-derivation mf:key-derivation-name="PBKDF2"
                    mf:salt="c2FsdA==" mf:iteration-count="2048"/>
                </mf:encryption-data>
              </mf:file-entry>
            </mf:manifest>
        "#;

        let entries = parse_manifest(xml).expect("manifest");
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].path, "content.xml");
        assert_eq!(entries[0].media_type.as_deref(), Some("text/xml"));

        let encryption = format_odf_encryption_metadata(&entries[0]).expect("encryption");
        assert!(encryption.contains("aes256-cbc"));
        assert!(encryption.contains("PBKDF2"));
        assert!(encryption.contains("2048"));
    }

    #[test]
    fn parse_manifest_reports_malformed_file_entry_attributes() {
        let xml = r#"
            <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
              <manifest:file-entry manifest:full-path="content.xml" manifest:full-path="meta.xml"/>
            </manifest:manifest>
        "#;

        let err = parse_manifest(xml).expect_err("malformed manifest attributes must fail");
        match err {
            crate::error::ParseError::Xml { file, .. } => {
                assert_eq!(file, "META-INF/manifest.xml");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_manifest_reports_malformed_encryption_number_attributes() {
        for xml in [
            r#"
            <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
              <manifest:file-entry manifest:full-path="content.xml">
                <manifest:encryption-data>
                  <manifest:algorithm manifest:key-size="bad"/>
                </manifest:encryption-data>
              </manifest:file-entry>
            </manifest:manifest>
        "#,
            r#"
            <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
              <manifest:file-entry manifest:full-path="content.xml">
                <manifest:encryption-data>
                  <manifest:key-derivation manifest:iteration-count="bad"/>
                </manifest:encryption-data>
              </manifest:file-entry>
            </manifest:manifest>
        "#,
        ] {
            let err = parse_manifest(xml).expect_err("malformed manifest number must fail");
            match err {
                crate::error::ParseError::Xml { file, .. } => {
                    assert_eq!(file, "META-INF/manifest.xml");
                }
                other => panic!("unexpected error: {other:?}"),
            }
        }
    }
}
