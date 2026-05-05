use super::{
    is_manifest_entry_encrypted, parse_content, parse_manifest, parse_styles, spreadsheet,
    Diagnostics, Document, DocumentFormat, IRNode, IrStore, OdfAtomicLimits, OdfEncryptionData,
    OdfLimits, OdfManifestEntry, OdfParser, ParseError, ParserConfig, SecureZipReader,
};
use crate::diagnostics::{push_info, push_warning};
use crate::xml_utils::{local_name, scan_xml_events, XmlScanControl};
use aes::{Aes128, Aes256};
use cbc::cipher::{block_padding::Pkcs7, BlockDecryptMut, KeyIvInit};
use cbc::Decryptor;
use docir_core::ir::DocumentMetadata;
use pbkdf2::pbkdf2_hmac;
use quick_xml::events::Event;
use quick_xml::Reader;
use sha1::Sha1;
use std::io::{Read, Seek};
use std::sync::Arc;

type StylesSettingsSignatures = (Option<String>, Option<String>, Option<String>);

impl OdfParser {
    pub(super) fn load_mimetype_and_manifest<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
    ) -> Result<(DocumentFormat, Vec<OdfManifestEntry>), ParseError> {
        let mimetype = zip
            .read_file_string("mimetype")
            .map(|s| s.trim().to_string())
            .map_err(|_| ParseError::UnsupportedFormat("Missing ODF mimetype".to_string()))?;

        let format = detect_odf_format(&mimetype).ok_or_else(|| {
            ParseError::UnsupportedFormat(format!("Unsupported ODF mimetype: {mimetype}"))
        })?;

        let manifest_entries = if zip.contains("META-INF/manifest.xml") {
            let manifest_xml = zip.read_file_string("META-INF/manifest.xml")?;
            parse_manifest(&manifest_xml)?
        } else {
            Vec::new()
        };

        Ok((format, manifest_entries))
    }

    pub(super) fn load_styles_settings_signatures<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
        doc: &mut Document,
    ) -> Result<StylesSettingsSignatures, ParseError> {
        let mut styles_xml: Option<String> = None;
        if zip.contains("styles.xml") {
            let xml = zip.read_file_string("styles.xml")?;
            if let Some(styles) = parse_styles(&xml) {
                let style_id = styles.id;
                store.insert(IRNode::StyleSet(styles));
                doc.styles = Some(style_id);
            }
            styles_xml = Some(xml);
        }

        let settings_xml = if zip.contains("settings.xml") {
            Some(zip.read_file_string("settings.xml")?)
        } else {
            None
        };

        let signatures_xml = if zip.contains("META-INF/documentsignatures.xml") {
            Some(zip.read_file_string("META-INF/documentsignatures.xml")?)
        } else {
            None
        };

        Ok((styles_xml, settings_xml, signatures_xml))
    }
}

pub(super) struct ContentState {
    pub(super) content_xml: Option<String>,
    pub(super) content_bytes: Option<Vec<u8>>,
    pub(super) fast_mode: bool,
    pub(super) content_size: Option<u64>,
}

pub(super) fn load_meta<R: Read + Seek>(
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
    doc: &mut Document,
) -> Result<(), ParseError> {
    if zip.contains("meta.xml") {
        let meta_xml = zip.read_file_string("meta.xml")?;
        if let Some(meta) = parse_meta(&meta_xml) {
            let meta_id = meta.id;
            store.insert(IRNode::Metadata(meta));
            doc.metadata = Some(meta_id);
        }
    }
    Ok(())
}

pub(super) fn handle_content_xml<R: Read + Seek>(
    config: &ParserConfig,
    zip: &mut SecureZipReader<R>,
    format: DocumentFormat,
    manifest_entries: &[OdfManifestEntry],
    store: &mut IrStore,
    doc: &mut Document,
    diagnostics: &mut Diagnostics,
) -> Result<ContentState, ParseError> {
    let mut content_xml: Option<String> = None;
    let mut content_bytes: Option<Vec<u8>> = None;
    let content_entry = manifest_entries
        .iter()
        .find(|entry| entry.path == "content.xml");
    let (content_size, fast_mode) = determine_content_mode(config, zip, format)?;

    if content_size.is_some() {
        let xml_bytes = read_content_bytes(config, zip, content_entry, diagnostics)?;
        if !xml_bytes.is_empty() {
            if !fast_mode {
                content_xml = Some(String::from_utf8_lossy(&xml_bytes).to_string());
            }
            parse_and_attach_content(config, format, fast_mode, &xml_bytes, store, doc)?;
        }
        content_bytes = Some(xml_bytes);
    }

    Ok(ContentState {
        content_xml,
        content_bytes,
        fast_mode,
        content_size,
    })
}

fn determine_content_mode<R: Read + Seek>(
    config: &ParserConfig,
    zip: &mut SecureZipReader<R>,
    format: DocumentFormat,
) -> Result<(Option<u64>, bool), ParseError> {
    if !zip.contains("content.xml") {
        return Ok((None, false));
    }

    let size = zip.file_size("content.xml")?;
    if let Some(max_bytes) = config.odf.max_bytes {
        if size > max_bytes {
            return Err(ParseError::ResourceLimit(format!(
                "ODF content.xml too large: {} bytes (max: {} bytes)",
                size, max_bytes
            )));
        }
    }

    let fast_mode = format == DocumentFormat::OdfSpreadsheet
        && (config.odf.force_fast || size >= config.odf.fast_threshold_bytes);
    Ok((Some(size), fast_mode))
}

fn read_content_bytes<R: Read + Seek>(
    config: &ParserConfig,
    zip: &mut SecureZipReader<R>,
    content_entry: Option<&OdfManifestEntry>,
    diagnostics: &mut Diagnostics,
) -> Result<Vec<u8>, ParseError> {
    let content_encrypted = content_entry
        .map(is_manifest_entry_encrypted)
        .unwrap_or(false);
    if !content_encrypted {
        return zip.read_file("content.xml");
    }

    let password = config.odf.password.as_deref();
    let encryption = content_entry.and_then(|entry| entry.encryption.as_ref());
    if let (Some(password), Some(encryption)) = (password, encryption) {
        match decrypt_odf_part(zip.read_file("content.xml")?, encryption, password) {
            Ok(bytes) => {
                push_info(
                    diagnostics,
                    "ODF_DECRYPT_OK",
                    "ODF encrypted content.xml decrypted successfully".to_string(),
                    Some("content.xml"),
                );
                Ok(bytes)
            }
            Err(message) => Err(ParseError::InvalidFormat(format!(
                "ODF decryption failed: {}",
                message
            ))),
        }
    } else {
        push_warning(
            diagnostics,
            "ODF_DECRYPT_SKIPPED",
            "ODF content.xml is encrypted but no password or encryption data is available"
                .to_string(),
            Some("content.xml"),
        );
        Ok(Vec::new())
    }
}

fn parse_and_attach_content(
    config: &ParserConfig,
    format: DocumentFormat,
    fast_mode: bool,
    xml_bytes: &[u8],
    store: &mut IrStore,
    doc: &mut Document,
) -> Result<(), ParseError> {
    let use_parallel =
        format == DocumentFormat::OdfSpreadsheet && config.odf.parallel_sheets && !fast_mode;
    let content_result = if use_parallel {
        let limits = Arc::new(OdfAtomicLimits::new(config, fast_mode));
        spreadsheet::parse_content_spreadsheet_parallel(xml_bytes, store, &limits, config)?
    } else {
        let limits = OdfLimits::new(config, fast_mode);
        parse_content(xml_bytes, format, store, &limits)?
    };

    doc.content.extend(content_result.content);
    doc.comments.extend(content_result.comments);
    doc.footnotes.extend(content_result.footnotes);
    doc.endnotes.extend(content_result.endnotes);
    doc.pivot_caches.extend(content_result.pivot_caches);
    Ok(())
}

fn detect_odf_format(mimetype: &str) -> Option<DocumentFormat> {
    let lower = mimetype.to_ascii_lowercase();
    if lower.contains("opendocument.text") || lower.contains("vnd.sun.xml.writer") {
        Some(DocumentFormat::OdfText)
    } else if lower.contains("opendocument.spreadsheet") || lower.contains("vnd.sun.xml.calc") {
        Some(DocumentFormat::OdfSpreadsheet)
    } else if lower.contains("opendocument.presentation") || lower.contains("vnd.sun.xml.impress") {
        Some(DocumentFormat::OdfPresentation)
    } else {
        None
    }
}

fn decrypt_odf_part(
    encrypted: Vec<u8>,
    encryption: &OdfEncryptionData,
    password: &str,
) -> Result<Vec<u8>, String> {
    let algorithm = encryption
        .algorithm_name
        .as_deref()
        .unwrap_or("http://www.w3.org/2001/04/xmlenc#aes256-cbc");
    let salt = encryption
        .salt
        .as_ref()
        .ok_or_else(|| "Missing encryption salt".to_string())?;
    let iv = encryption
        .init_vector
        .as_ref()
        .ok_or_else(|| "Missing encryption IV".to_string())?;
    let iterations = encryption.iteration_count.unwrap_or(100_000);
    let key_bits = encryption
        .key_size
        .or_else(|| {
            if algorithm.contains("aes256") {
                Some(256)
            } else if algorithm.contains("aes128") {
                Some(128)
            } else {
                None
            }
        })
        .ok_or_else(|| "Unsupported encryption algorithm".to_string())?;
    let key_len = (key_bits / 8) as usize;
    if iv.len() != 16 {
        return Err(format!("Unsupported IV length: {}", iv.len()));
    }

    let mut key = vec![0u8; key_len];
    // Security note: PBKDF2 with SHA-1 is required by the ODF specification
    // (OpenDocument 1.3, section 4.4). SHA-1 is deprecated for cryptographic
    // use but must be used here for spec compliance.
    pbkdf2_hmac::<Sha1>(password.as_bytes(), salt, iterations, &mut key);

    let mut buffer = encrypted;
    if key_len == 32 {
        let decryptor = Decryptor::<Aes256>::new_from_slices(&key, iv)
            .map_err(|_| "Invalid AES-256 key or IV".to_string())?;
        let decrypted = decryptor
            .decrypt_padded_mut::<Pkcs7>(&mut buffer)
            .map_err(|_| "Invalid AES-256 padding".to_string())?;
        Ok(decrypted.to_vec())
    } else if key_len == 16 {
        let decryptor = Decryptor::<Aes128>::new_from_slices(&key, iv)
            .map_err(|_| "Invalid AES-128 key or IV".to_string())?;
        let decrypted = decryptor
            .decrypt_padded_mut::<Pkcs7>(&mut buffer)
            .map_err(|_| "Invalid AES-128 padding".to_string())?;
        Ok(decrypted.to_vec())
    } else {
        Err(format!("Unsupported key length: {}", key_len))
    }
}

fn parse_meta(xml: &str) -> Option<DocumentMetadata> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut meta = DocumentMetadata::new();
    let mut current = None;

    #[derive(Clone, Copy)]
    enum MetaField {
        Title,
        Subject,
        Creator,
        Keywords,
        Description,
        Created,
        Modified,
    }

    if scan_xml_events(&mut reader, &mut buf, "meta.xml", |event| {
        match event {
            Event::Start(e) => {
                current = match local_name(e.name().as_ref()) {
                    b"title" => Some(MetaField::Title),
                    b"subject" => Some(MetaField::Subject),
                    b"creator" => Some(MetaField::Creator),
                    b"keyword" => Some(MetaField::Keywords),
                    b"description" => Some(MetaField::Description),
                    b"creation-date" => Some(MetaField::Created),
                    b"date" => Some(MetaField::Modified),
                    _ => None,
                };
            }
            Event::Text(e) => {
                if let Some(field) = current {
                    let value = e.unescape().unwrap_or_default().to_string();
                    match field {
                        MetaField::Title => meta.title = Some(value),
                        MetaField::Subject => meta.subject = Some(value),
                        MetaField::Creator => meta.creator = Some(value),
                        MetaField::Keywords => meta.keywords = Some(value),
                        MetaField::Description => meta.description = Some(value),
                        MetaField::Created => meta.created = Some(value),
                        MetaField::Modified => meta.modified = Some(value),
                    }
                }
            }
            Event::End(_) => {
                current = None;
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })
    .is_err()
    {
        return None;
    }

    let has_any = meta.title.is_some()
        || meta.subject.is_some()
        || meta.creator.is_some()
        || meta.keywords.is_some()
        || meta.description.is_some()
        || meta.created.is_some()
        || meta.modified.is_some();

    if has_any {
        Some(meta)
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn encryption_data() -> OdfEncryptionData {
        OdfEncryptionData {
            checksum_type: None,
            checksum: None,
            algorithm_name: Some("http://www.w3.org/2001/04/xmlenc#aes256-cbc".to_string()),
            init_vector: Some(vec![0_u8; 16]),
            key_derivation_name: None,
            salt: Some(vec![1_u8; 16]),
            iteration_count: Some(10),
            key_size: Some(256),
        }
    }

    #[test]
    fn detect_odf_format_supports_expected_mimetypes() {
        assert_eq!(
            detect_odf_format("application/vnd.oasis.opendocument.text"),
            Some(DocumentFormat::OdfText)
        );
        assert_eq!(
            detect_odf_format("application/vnd.sun.xml.calc"),
            Some(DocumentFormat::OdfSpreadsheet)
        );
        assert_eq!(
            detect_odf_format("application/vnd.oasis.opendocument.presentation"),
            Some(DocumentFormat::OdfPresentation)
        );
        assert_eq!(detect_odf_format("application/octet-stream"), None);
    }

    #[test]
    fn parse_meta_extracts_known_fields_and_handles_empty_or_malformed_xml() {
        let meta = parse_meta(
            r#"
            <office:meta xmlns:dc="dc" xmlns:meta="meta">
              <dc:title>Title</dc:title>
              <dc:subject>Subject</dc:subject>
              <dc:creator>Alice</dc:creator>
              <meta:keyword>tag1,tag2</meta:keyword>
              <dc:description>Desc</dc:description>
              <meta:creation-date>2026-01-01</meta:creation-date>
              <dc:date>2026-01-02</dc:date>
            </office:meta>
            "#,
        )
        .expect("meta must parse");
        assert_eq!(meta.title.as_deref(), Some("Title"));
        assert_eq!(meta.subject.as_deref(), Some("Subject"));
        assert_eq!(meta.creator.as_deref(), Some("Alice"));
        assert_eq!(meta.keywords.as_deref(), Some("tag1,tag2"));
        assert_eq!(meta.description.as_deref(), Some("Desc"));
        assert_eq!(meta.created.as_deref(), Some("2026-01-01"));
        assert_eq!(meta.modified.as_deref(), Some("2026-01-02"));

        let prefixed_meta = parse_meta(
            r#"
            <pkg:meta xmlns:dct="dc" xmlns:m="meta">
              <dct:title>Alt Title</dct:title>
              <dct:creator>Bob</dct:creator>
              <m:creation-date>2026-02-01</m:creation-date>
            </pkg:meta>
            "#,
        )
        .expect("alternate-prefix meta must parse");
        assert_eq!(prefixed_meta.title.as_deref(), Some("Alt Title"));
        assert_eq!(prefixed_meta.creator.as_deref(), Some("Bob"));
        assert_eq!(prefixed_meta.created.as_deref(), Some("2026-02-01"));

        assert!(parse_meta("<office:meta/>").is_none());
        assert!(parse_meta("<office:meta><dc:title>").is_none());
    }

    #[test]
    fn decrypt_odf_part_validates_required_encryption_fields() {
        let mut enc = encryption_data();
        enc.salt = None;
        let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("missing salt");
        assert!(err.contains("Missing encryption salt"));

        let mut enc = encryption_data();
        enc.init_vector = None;
        let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("missing iv");
        assert!(err.contains("Missing encryption IV"));

        let mut enc = encryption_data();
        enc.init_vector = Some(vec![0_u8; 8]);
        let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("bad iv length");
        assert!(err.contains("Unsupported IV length: 8"));
    }

    #[test]
    fn decrypt_odf_part_rejects_unsupported_algorithm_or_key_length() {
        let mut enc = encryption_data();
        enc.algorithm_name = Some("urn:unknown".to_string());
        enc.key_size = None;
        let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("unsupported algo");
        assert!(err.contains("Unsupported encryption algorithm"));

        let mut enc = encryption_data();
        enc.key_size = Some(192);
        let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("unsupported key size");
        assert!(err.contains("Unsupported key length: 24"));
    }
}
