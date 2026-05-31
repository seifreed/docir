use std::io::{Read, Seek};
use std::sync::Arc;

use super::{
    Diagnostics, Document, DocumentFormat, IrStore, OdfAtomicLimits, OdfLimits, OdfManifestEntry,
    ParseError, ParserConfig, SecureZipReader, is_manifest_entry_encrypted, parse_content,
    spreadsheet,
};
use crate::diagnostics::{push_info, push_warning};

use super::container_encryption::decrypt_odf_part;

pub(super) struct ContentState {
    pub(super) content_xml: Option<String>,
    pub(super) content_bytes: Option<Vec<u8>>,
    pub(super) fast_mode: bool,
    pub(super) content_size: Option<u64>,
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
    if let Some(max_bytes) = config.odf.max_bytes
        && size > max_bytes
    {
        return Err(ParseError::ResourceLimit(format!(
            "ODF content.xml too large: {} bytes (max: {} bytes)",
            size, max_bytes
        )));
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
