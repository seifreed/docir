use crate::ParsedDocument;
use docir_core::types::DocumentFormat;
use docir_core::{ExtractedArtifact, ExtractedArtifactKind, ExtractionManifest, ExtractionWarning};
use docir_parser::ole::{is_ole_container, Cfb};
use docir_parser::zip_handler::{SecureZipReader, ZipConfig};
use sha2::{Digest, Sha256};
use std::collections::HashSet;
use std::io::Cursor;

/// Runtime options controlling artifact extraction outputs.
#[derive(Debug, Clone)]
pub struct ArtifactExtractionOptions {
    pub compute_hashes: bool,
    pub with_raw: bool,
    pub no_media: bool,
    pub only_ole: bool,
    pub only_rtf_objects: bool,
}

impl Default for ArtifactExtractionOptions {
    fn default() -> Self {
        Self {
            compute_hashes: true,
            with_raw: false,
            no_media: false,
            only_ole: false,
            only_rtf_objects: false,
        }
    }
}

/// A binary payload ready to be written by adapters such as the CLI.
#[derive(Debug, Clone)]
pub struct ExtractedPayload {
    pub artifact_id: String,
    pub relative_path: String,
    pub data: Vec<u8>,
}

/// In-memory extraction result consumed by CLI and bindings.
#[derive(Debug, Clone, Default)]
pub struct ArtifactExtractionBundle {
    pub manifest: ExtractionManifest,
    pub payloads: Vec<ExtractedPayload>,
}

#[derive(Debug, Clone)]
struct EmbeddedPayload {
    stream_name: String,
    file_name: Option<String>,
    source_path: Option<String>,
    temp_path: Option<String>,
    data: Vec<u8>,
}

/// Extracts embedded artifacts from the original container bytes.
pub fn extract_artifacts_from_bytes(
    parsed: &ParsedDocument,
    input_bytes: &[u8],
    source_document: Option<String>,
    zip_config: &ZipConfig,
    options: &ArtifactExtractionOptions,
) -> ArtifactExtractionBundle {
    let mut bundle = ArtifactExtractionBundle {
        manifest: ExtractionManifest::new(),
        ..ArtifactExtractionBundle::default()
    };
    bundle.manifest.source_document = source_document;

    match parsed.format() {
        DocumentFormat::WordProcessing
        | DocumentFormat::Spreadsheet
        | DocumentFormat::Presentation => {
            if is_legacy_cfb_document(parsed) {
                if options.only_rtf_objects {
                    bundle.manifest.warnings.push(ExtractionWarning::new(
                        "NO_MATCHING_ARTIFACTS",
                        "RTF-only extraction requested for a legacy CFB container",
                    ));
                    return bundle;
                }
                extract_legacy_cfb_artifacts(input_bytes, options, &mut bundle);
                return bundle;
            }
            if options.only_rtf_objects {
                bundle.manifest.warnings.push(ExtractionWarning::new(
                    "NO_MATCHING_ARTIFACTS",
                    "RTF-only extraction requested for a non-RTF container",
                ));
                return bundle;
            }
            extract_ooxml_artifacts(input_bytes, zip_config, options, &mut bundle);
        }
        DocumentFormat::Rtf => {
            extract_rtf_artifacts(input_bytes, options, &mut bundle);
        }
        _ => {
            bundle.manifest.warnings.push(ExtractionWarning::new(
                "UNSUPPORTED_EXTRACTION_FORMAT",
                format!(
                    "Embedded artifact extraction is not implemented for {}",
                    parsed.format().extension()
                ),
            ));
        }
    }

    if bundle.manifest.artifacts.is_empty() {
        bundle.manifest.warnings.push(ExtractionWarning::new(
            "NO_ARTIFACTS",
            "No extractable embedded artifacts were found",
        ));
    }

    bundle
}

fn extract_ooxml_artifacts(
    input_bytes: &[u8],
    zip_config: &ZipConfig,
    options: &ArtifactExtractionOptions,
    bundle: &mut ArtifactExtractionBundle,
) {
    let Ok(mut zip) = SecureZipReader::new(Cursor::new(input_bytes), zip_config.clone()) else {
        bundle.manifest.warnings.push(ExtractionWarning::new(
            "ZIP_OPEN_FAILED",
            "Unable to reopen the OOXML package for artifact extraction",
        ));
        return;
    };

    let mut seen = HashSet::new();
    let mut paths: Vec<(String, ExtractedArtifactKind)> = zip
        .list_prefix("word/embeddings/")
        .into_iter()
        .chain(zip.list_prefix("xl/embeddings/"))
        .chain(zip.list_prefix("ppt/embeddings/"))
        .filter(|p| p.ends_with(".bin") || p.ends_with(".ole"))
        .map(|p| (p.to_string(), ExtractedArtifactKind::OleObject))
        .collect();
    paths.extend(
        zip.list_prefix("word/media/")
            .into_iter()
            .chain(zip.list_prefix("xl/media/"))
            .chain(zip.list_prefix("ppt/media/"))
            .filter(|_| !options.no_media)
            .map(|p| (p.to_string(), ExtractedArtifactKind::MediaAsset)),
    );
    paths.extend(
        zip.list_prefix("word/activeX/")
            .into_iter()
            .chain(zip.list_prefix("xl/activeX/"))
            .chain(zip.list_prefix("ppt/activeX/"))
            .filter(|p| p.ends_with(".bin") || p.ends_with(".dat"))
            .map(|p| (p.to_string(), ExtractedArtifactKind::ActiveXControl)),
    );
    paths.sort_by(|left, right| left.0.cmp(&right.0));

    let mut ole_index = 0usize;
    let mut media_index = 0usize;
    let mut activex_index = 0usize;
    let mut payload_index = 0usize;

    for (path, kind) in paths {
        if !seen.insert(path.clone()) {
            continue;
        }
        if options.only_ole && kind != ExtractedArtifactKind::OleObject {
            continue;
        }

        let Ok(data) = zip.read_file(&path) else {
            bundle.manifest.warnings.push(ExtractionWarning::new(
                "ARTIFACT_READ_FAILED",
                format!("Unable to read embedded artifact {}", path),
            ));
            continue;
        };

        let ordinal = match kind {
            ExtractedArtifactKind::OleObject => {
                ole_index += 1;
                ole_index
            }
            ExtractedArtifactKind::MediaAsset => {
                media_index += 1;
                media_index
            }
            ExtractedArtifactKind::ActiveXControl => {
                activex_index += 1;
                activex_index
            }
            _ => 0,
        };
        let prefix = match kind {
            ExtractedArtifactKind::OleObject => "ole-object",
            ExtractedArtifactKind::MediaAsset => "media-asset",
            ExtractedArtifactKind::ActiveXControl => "activex-control",
            _ => "artifact",
        };

        let mut artifact = ExtractedArtifact::new(format!("{prefix}-{ordinal}"), kind);
        artifact.source_path = Some(path.clone());
        artifact.suggested_name = Some(file_name_from_path(&path));
        artifact.size_bytes = Some(data.len() as u64);
        assign_sha256(&mut artifact.sha256, &data, options.compute_hashes);
        let (_, mime_type) = if kind == ExtractedArtifactKind::MediaAsset {
            classify_media_asset(&path, &data)
        } else {
            classify_payload(&data, artifact.suggested_name.as_deref())
        };
        artifact.mime_type = Some(mime_type.to_string());

        if kind == ExtractedArtifactKind::MediaAsset {
            let file_name = artifact
                .suggested_name
                .clone()
                .unwrap_or_else(|| format!("artifact_{ordinal}"));
            let relative_path = format!("payloads/{}", sanitize_name(&file_name));
            artifact.output_path = Some(relative_path.clone());
            bundle.payloads.push(ExtractedPayload {
                artifact_id: artifact.id.clone(),
                relative_path,
                data: data.clone(),
            });
            bundle.manifest.artifacts.push(artifact);
            continue;
        }

        if options.with_raw {
            let raw_name = format!("{}_{}", prefix, sanitize_name(&path));
            let relative_path = format!("raw/{}", raw_name);
            artifact.output_path = Some(relative_path.clone());
            bundle.payloads.push(ExtractedPayload {
                artifact_id: artifact.id.clone(),
                relative_path,
                data: data.clone(),
            });
        }

        bundle.manifest.artifacts.push(artifact);

        if let Some(payload) = extract_embedded_payload(&data) {
            payload_index += 1;
            let mut payload_artifact = ExtractedArtifact::new(
                format!("ole-native-payload-{}", payload_index),
                ExtractedArtifactKind::OleNativePayload,
            );
            payload_artifact.source_path = Some(format!("{}#{}", path, payload.stream_name));
            payload_artifact.suggested_name = payload.file_name.clone();
            payload_artifact.size_bytes = Some(payload.data.len() as u64);
            assign_sha256(
                &mut payload_artifact.sha256,
                &payload.data,
                options.compute_hashes,
            );
            let (payload_kind, mime_type) =
                classify_payload(&payload.data, payload.file_name.as_deref());
            payload_artifact.mime_type = Some(mime_type.to_string());
            payload_artifact.encoding = None;
            let file_name = preferred_output_name(
                payload.file_name.as_deref(),
                payload_index,
                payload_kind,
                payload_artifact.mime_type.as_deref(),
            );
            let relative_path = format!("payloads/{}", file_name);
            payload_artifact.output_path = Some(relative_path.clone());
            if let Some(source_path) = payload.source_path {
                payload_artifact
                    .errors
                    .push(format!("source_path={source_path}"));
            }
            if let Some(temp_path) = payload.temp_path {
                payload_artifact
                    .errors
                    .push(format!("temp_path={temp_path}"));
            }
            bundle.payloads.push(ExtractedPayload {
                artifact_id: payload_artifact.id.clone(),
                relative_path,
                data: payload.data,
            });
            bundle.manifest.artifacts.push(payload_artifact);
        }
    }
}

fn is_legacy_cfb_document(parsed: &ParsedDocument) -> bool {
    parsed
        .document()
        .and_then(|doc| doc.span.as_ref())
        .map(|span| span.file_path.starts_with("cfb:/"))
        .unwrap_or(false)
}

fn extract_rtf_artifacts(
    input_bytes: &[u8],
    options: &ArtifactExtractionOptions,
    bundle: &mut ArtifactExtractionBundle,
) {
    let mut payload_index = 0usize;
    for (idx, blob) in scan_rtf_objdata(input_bytes).into_iter().enumerate() {
        if options.only_ole && !is_ole_container(&blob) {
            continue;
        }

        payload_index += 1;
        let mut artifact = ExtractedArtifact::new(
            format!("rtf-object-{}", idx + 1),
            ExtractedArtifactKind::RtfEmbeddedObject,
        );
        artifact.source_path = Some(format!("rtf:objdata#{}", idx + 1));
        artifact.size_bytes = Some(blob.len() as u64);
        assign_sha256(&mut artifact.sha256, &blob, options.compute_hashes);
        let (payload_kind, mime_type) = classify_payload(&blob, None);
        artifact.mime_type = Some(mime_type.to_string());
        let file_name = preferred_output_name(None, payload_index, payload_kind, Some(mime_type));
        let relative_path = format!("rtf/{}", file_name);
        artifact.output_path = Some(relative_path.clone());
        artifact.suggested_name = Some(file_name.clone());
        bundle.payloads.push(ExtractedPayload {
            artifact_id: artifact.id.clone(),
            relative_path,
            data: blob.clone(),
        });
        bundle.manifest.artifacts.push(artifact);

        if let Some(payload) = extract_embedded_payload(&blob) {
            payload_index += 1;
            let mut payload_artifact = ExtractedArtifact::new(
                format!("ole-native-payload-{}", payload_index),
                ExtractedArtifactKind::OleNativePayload,
            );
            payload_artifact.source_path =
                Some(format!("rtf:objdata#{}#{}", idx + 1, payload.stream_name));
            payload_artifact.size_bytes = Some(payload.data.len() as u64);
            assign_sha256(
                &mut payload_artifact.sha256,
                &payload.data,
                options.compute_hashes,
            );
            let (inner_kind, inner_mime) =
                classify_payload(&payload.data, payload.file_name.as_deref());
            payload_artifact.mime_type = Some(inner_mime.to_string());
            let inner_name = preferred_output_name(
                payload.file_name.as_deref(),
                payload_index,
                inner_kind,
                Some(inner_mime),
            );
            let inner_relative_path = format!("payloads/{}", inner_name);
            payload_artifact.output_path = Some(inner_relative_path.clone());
            payload_artifact.suggested_name = Some(inner_name);
            bundle.payloads.push(ExtractedPayload {
                artifact_id: payload_artifact.id.clone(),
                relative_path: inner_relative_path,
                data: payload.data,
            });
            bundle.manifest.artifacts.push(payload_artifact);
        }
    }
}

fn extract_legacy_cfb_artifacts(
    input_bytes: &[u8],
    options: &ArtifactExtractionOptions,
    bundle: &mut ArtifactExtractionBundle,
) {
    let Ok(cfb) = Cfb::parse(input_bytes.to_vec()) else {
        bundle.manifest.warnings.push(ExtractionWarning::new(
            "CFB_OPEN_FAILED",
            "Unable to open the legacy Office CFB container",
        ));
        return;
    };

    let mut paths = cfb.list_streams();
    paths.sort();
    let mut payload_index = 0usize;

    for path in paths {
        let metadata = cfb.entry_metadata(&path).cloned();
        let upper = path.to_ascii_uppercase();
        let is_ole_object = upper.contains("OBJECTPOOL/")
            || upper.ends_with("OLE10NATIVE")
            || upper == "PACKAGE"
            || upper.ends_with("/PACKAGE")
            || upper.ends_with("/CONTENTS");
        if !is_ole_object {
            continue;
        }
        if options.only_ole
            && !(upper.contains("OBJECTPOOL/")
                || upper.ends_with("OLE10NATIVE")
                || upper == "PACKAGE"
                || upper.ends_with("/PACKAGE"))
        {
            continue;
        }

        let Some(data) = cfb.read_stream(&path) else {
            continue;
        };

        payload_index += 1;
        let mut artifact = ExtractedArtifact::new(
            format!("legacy-ole-object-{}", payload_index),
            ExtractedArtifactKind::OleObject,
        );
        artifact.source_path = Some(path.clone());
        artifact.suggested_name = Some(file_name_from_path(&path));
        artifact.size_bytes = Some(data.len() as u64);
        artifact.start_sector = metadata.as_ref().map(|entry| entry.start_sector);
        artifact.created_filetime = metadata.as_ref().and_then(|entry| entry.created_filetime);
        artifact.modified_filetime = metadata.as_ref().and_then(|entry| entry.modified_filetime);
        assign_sha256(&mut artifact.sha256, &data, options.compute_hashes);
        let (_, mime_type) = classify_payload(&data, artifact.suggested_name.as_deref());
        artifact.mime_type = Some(mime_type.to_string());

        if options.with_raw {
            let raw_name = format!("legacy_{}", sanitize_name(&path));
            let relative_path = format!("raw/{}", raw_name);
            artifact.output_path = Some(relative_path.clone());
            bundle.payloads.push(ExtractedPayload {
                artifact_id: artifact.id.clone(),
                relative_path,
                data: data.clone(),
            });
        }

        bundle.manifest.artifacts.push(artifact);

        let payload = if upper.ends_with("OLE10NATIVE") {
            parse_ole10_native(&data).map(|payload| EmbeddedPayload {
                stream_name: path.clone(),
                file_name: payload.file_name,
                source_path: payload.source_path,
                temp_path: payload.temp_path,
                data: payload.data,
            })
        } else if upper.ends_with("/PACKAGE") || upper == "PACKAGE" {
            Some(EmbeddedPayload {
                stream_name: path.clone(),
                file_name: None,
                source_path: None,
                temp_path: None,
                data,
            })
        } else {
            extract_embedded_payload(&data)
        };

        if let Some(payload) = payload {
            payload_index += 1;
            let mut payload_artifact = ExtractedArtifact::new(
                format!("legacy-payload-{}", payload_index),
                ExtractedArtifactKind::OleNativePayload,
            );
            payload_artifact.source_path = Some(format!("{}#{}", path, payload.stream_name));
            payload_artifact.suggested_name = payload.file_name.clone();
            payload_artifact.size_bytes = Some(payload.data.len() as u64);
            payload_artifact.start_sector = metadata.as_ref().map(|entry| entry.start_sector);
            payload_artifact.created_filetime =
                metadata.as_ref().and_then(|entry| entry.created_filetime);
            payload_artifact.modified_filetime =
                metadata.as_ref().and_then(|entry| entry.modified_filetime);
            assign_sha256(
                &mut payload_artifact.sha256,
                &payload.data,
                options.compute_hashes,
            );
            let (kind, mime_type) = classify_payload(&payload.data, payload.file_name.as_deref());
            payload_artifact.mime_type = Some(mime_type.to_string());
            let file_name = preferred_output_name(
                payload.file_name.as_deref(),
                payload_index,
                kind,
                payload_artifact.mime_type.as_deref(),
            );
            let relative_path = format!("payloads/{}", file_name);
            payload_artifact.output_path = Some(relative_path.clone());
            bundle.payloads.push(ExtractedPayload {
                artifact_id: payload_artifact.id.clone(),
                relative_path,
                data: payload.data,
            });
            bundle.manifest.artifacts.push(payload_artifact);
        }
    }
}

fn extract_embedded_payload(data: &[u8]) -> Option<EmbeddedPayload> {
    if !is_ole_container(data) {
        return None;
    }
    let cfb = Cfb::parse(data.to_vec()).ok()?;
    for stream_name in ["\u{0001}Ole10Native", "Ole10Native", "Package"] {
        let Some(stream) = cfb.read_stream(stream_name) else {
            continue;
        };
        if stream_name.contains("Ole10Native") {
            if let Some(payload) = parse_ole10_native(&stream) {
                return Some(EmbeddedPayload {
                    stream_name: stream_name.to_string(),
                    file_name: payload.file_name,
                    source_path: payload.source_path,
                    temp_path: payload.temp_path,
                    data: payload.data,
                });
            }
        } else {
            return Some(EmbeddedPayload {
                stream_name: stream_name.to_string(),
                file_name: None,
                source_path: None,
                temp_path: None,
                data: stream,
            });
        }
    }
    None
}

#[derive(Debug, Clone)]
struct Ole10NativePayload {
    file_name: Option<String>,
    source_path: Option<String>,
    temp_path: Option<String>,
    data: Vec<u8>,
}

fn parse_ole10_native(data: &[u8]) -> Option<Ole10NativePayload> {
    if data.len() < 6 {
        return None;
    }

    let mut offset = 4usize;
    offset = offset.checked_add(2)?;
    let file_name = read_c_string(data, &mut offset)?;
    let source_path = read_c_string(data, &mut offset)?;
    offset = offset.checked_add(8)?;
    let temp_path = read_c_string(data, &mut offset)?;
    if offset + 4 > data.len() {
        return None;
    }
    let size = u32::from_le_bytes(data[offset..offset + 4].try_into().ok()?) as usize;
    offset += 4;
    if offset + size > data.len() {
        return None;
    }

    Some(Ole10NativePayload {
        file_name: empty_to_none(file_name),
        source_path: empty_to_none(source_path),
        temp_path: empty_to_none(temp_path),
        data: data[offset..offset + size].to_vec(),
    })
}

fn read_c_string(data: &[u8], offset: &mut usize) -> Option<String> {
    let start = *offset;
    while *offset < data.len() && data[*offset] != 0 {
        *offset += 1;
    }
    if *offset >= data.len() {
        return None;
    }
    let value = String::from_utf8_lossy(&data[start..*offset]).to_string();
    *offset += 1;
    Some(value)
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

const MAX_HEX_BLOB_SIZE: usize = 100 * 1024 * 1024;

fn decode_hex_blob(hex: &[u8]) -> Option<Vec<u8>> {
    if hex.len() < 2 || hex.len() > MAX_HEX_BLOB_SIZE {
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

fn classify_payload(data: &[u8], file_name: Option<&str>) -> (&'static str, &'static str) {
    if data.starts_with(b"MZ") {
        return ("pe", "application/vnd.microsoft.portable-executable");
    }
    if data.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]) {
        return ("office", "application/x-cfb");
    }
    if data.starts_with(b"PK\x03\x04") {
        if file_name
            .map(|name| {
                let lower = name.to_ascii_lowercase();
                lower.ends_with(".docx")
                    || lower.ends_with(".docm")
                    || lower.ends_with(".xlsx")
                    || lower.ends_with(".xlsm")
                    || lower.ends_with(".pptx")
                    || lower.ends_with(".pptm")
            })
            .unwrap_or(false)
        {
            return ("office", "application/vnd.openxmlformats-officedocument");
        }
        return ("zip", "application/zip");
    }
    if data.starts_with(b"%PDF-") {
        return ("pdf", "application/pdf");
    }
    if looks_like_text(data) {
        return ("script", "text/plain");
    }
    ("unknown", "application/octet-stream")
}

fn classify_media_asset(path: &str, data: &[u8]) -> (&'static str, &'static str) {
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".png") {
        return ("image", "image/png");
    }
    if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        return ("image", "image/jpeg");
    }
    if lower.ends_with(".gif") {
        return ("image", "image/gif");
    }
    if lower.ends_with(".bmp") {
        return ("image", "image/bmp");
    }
    if lower.ends_with(".tif") || lower.ends_with(".tiff") {
        return ("image", "image/tiff");
    }
    if lower.ends_with(".webp") {
        return ("image", "image/webp");
    }
    if lower.ends_with(".wav") {
        return ("audio", "audio/wav");
    }
    if lower.ends_with(".mp3") {
        return ("audio", "audio/mpeg");
    }
    if lower.ends_with(".mp4") {
        return ("video", "video/mp4");
    }
    classify_payload(data, Some(path))
}

fn looks_like_text(data: &[u8]) -> bool {
    if data.is_empty() {
        return false;
    }
    let sample = &data[..data.len().min(256)];
    sample
        .iter()
        .all(|byte| byte.is_ascii_whitespace() || byte.is_ascii_graphic())
}

fn preferred_output_name(
    suggested_name: Option<&str>,
    index: usize,
    kind: &str,
    mime_type: Option<&str>,
) -> String {
    if let Some(name) = suggested_name.and_then(empty_to_none_str) {
        return sanitize_name(name);
    }
    let ext = match kind {
        "pe" => "exe",
        "pdf" => "pdf",
        "zip" => "zip",
        "office" => "bin",
        "script" => "txt",
        _ => mime_type
            .and_then(|value| value.rsplit('/').next())
            .unwrap_or("bin"),
    };
    format!("artifact_{}.{}", index, ext)
}

fn sanitize_name(input: &str) -> String {
    let mut out = String::with_capacity(input.len());
    for ch in input.chars() {
        if ch.is_ascii_alphanumeric() || ch == '.' || ch == '-' || ch == '_' {
            out.push(ch);
        } else {
            out.push('_');
        }
    }
    if out.is_empty() {
        "artifact.bin".to_string()
    } else {
        out
    }
}

fn file_name_from_path(path: &str) -> String {
    path.rsplit('/').next().unwrap_or(path).to_string()
}

fn sha256_hex(data: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(data);
    format!("{:x}", hasher.finalize())
}

fn assign_sha256(target: &mut Option<String>, data: &[u8], compute_hashes: bool) {
    if compute_hashes {
        *target = Some(sha256_hex(data));
    }
}

fn empty_to_none(value: String) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

fn empty_to_none_str(value: &str) -> Option<&str> {
    if value.trim().is_empty() {
        None
    } else {
        Some(value)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{test_support::build_test_cfb, DocirApp, ParserConfig};
    use std::io::Write;
    use zip::write::SimpleFileOptions;

    fn build_minimal_docx(extra_entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut cursor = Cursor::new(Vec::<u8>::new());
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            let options = SimpleFileOptions::default();
            writer
                .start_file("[Content_Types].xml", options)
                .expect("content types");
            writer
                .write_all(
                    br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
                )
                .expect("write content types");
            writer.start_file("_rels/.rels", options).expect("rels");
            writer
                .write_all(
                    br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
                )
                .expect("write rels");
            writer
                .start_file("word/document.xml", options)
                .expect("document");
            writer
                .write_all(
                    br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>docir</w:t></w:r></w:p></w:body>
</w:document>"#,
                )
                .expect("write document");
            for (path, data) in extra_entries {
                writer.start_file(path, options).expect("extra entry");
                writer.write_all(data).expect("extra data");
            }
            writer.finish().expect("finish zip");
        }
        cursor.into_inner()
    }

    #[test]
    fn parse_ole10_native_extracts_metadata_and_payload() {
        let mut blob = Vec::new();
        blob.extend_from_slice(&100u32.to_le_bytes());
        blob.extend_from_slice(&2u16.to_le_bytes());
        blob.extend_from_slice(b"payload.exe\0");
        blob.extend_from_slice(b"C:\\src\\payload.exe\0");
        blob.extend_from_slice(&0u32.to_le_bytes());
        blob.extend_from_slice(&0u32.to_le_bytes());
        blob.extend_from_slice(b"C:\\temp\\payload.exe\0");
        blob.extend_from_slice(&4u32.to_le_bytes());
        blob.extend_from_slice(b"MZ!!");

        let payload = parse_ole10_native(&blob).expect("payload");
        assert_eq!(payload.file_name.as_deref(), Some("payload.exe"));
        assert_eq!(payload.source_path.as_deref(), Some("C:\\src\\payload.exe"));
        assert_eq!(payload.temp_path.as_deref(), Some("C:\\temp\\payload.exe"));
        assert_eq!(payload.data, b"MZ!!");
    }

    #[test]
    fn scan_rtf_objdata_decodes_embedded_hex() {
        let blobs = scan_rtf_objdata(br"{\rtf1{\object{\objdata 4d5a9000}}}");
        assert_eq!(blobs, vec![vec![0x4d, 0x5a, 0x90, 0x00]]);
    }

    #[test]
    fn extract_artifacts_finds_ooxml_embedding_and_payload() {
        let bytes = build_minimal_docx(&[("word/embeddings/object1.bin", b"MZPAYLOAD")]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app.parse_bytes(&bytes).expect("parse bytes");
        let bundle = extract_artifacts_from_bytes(
            &parsed,
            &bytes,
            Some("memory.docx".to_string()),
            &ParserConfig::default().zip_config,
            &ArtifactExtractionOptions {
                with_raw: true,
                ..ArtifactExtractionOptions::default()
            },
        );

        assert!(bundle
            .manifest
            .artifacts
            .iter()
            .any(|artifact| artifact.kind == ExtractedArtifactKind::OleObject));
        assert!(bundle
            .payloads
            .iter()
            .any(|payload| payload.relative_path.starts_with("raw/")));
    }

    #[test]
    fn extract_artifacts_finds_legacy_cfb_payload() {
        let mut ole10 = Vec::new();
        ole10.extend_from_slice(&64u32.to_le_bytes());
        ole10.extend_from_slice(&2u16.to_le_bytes());
        ole10.extend_from_slice(b"dropper.exe\0");
        ole10.extend_from_slice(b"C:\\src\\dropper.exe\0");
        ole10.extend_from_slice(&0u32.to_le_bytes());
        ole10.extend_from_slice(&0u32.to_le_bytes());
        ole10.extend_from_slice(b"C:\\temp\\dropper.exe\0");
        ole10.extend_from_slice(&4u32.to_le_bytes());
        ole10.extend_from_slice(b"MZ!!");

        let bytes = build_test_cfb(&[("WordDocument", b"doc"), ("Ole10Native", &ole10)]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app.parse_bytes(&bytes).expect("parse legacy doc");
        let bundle = extract_artifacts_from_bytes(
            &parsed,
            &bytes,
            Some("legacy.doc".to_string()),
            &ParserConfig::default().zip_config,
            &ArtifactExtractionOptions::default(),
        );

        assert!(bundle
            .manifest
            .artifacts
            .iter()
            .any(|artifact| artifact.kind == ExtractedArtifactKind::OleObject));
        assert!(bundle
            .manifest
            .artifacts
            .iter()
            .any(|artifact| artifact.kind == ExtractedArtifactKind::OleObject
                && artifact.start_sector.is_some()));
        assert!(bundle
            .manifest
            .artifacts
            .iter()
            .any(|artifact| artifact.kind == ExtractedArtifactKind::OleNativePayload));
    }

    #[test]
    fn extract_artifacts_finds_legacy_package_payload() {
        let bytes = build_test_cfb(&[("Workbook", b"wb"), ("Package", b"%PDF-1.7")]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app.parse_bytes(&bytes).expect("parse legacy xls");
        let bundle = extract_artifacts_from_bytes(
            &parsed,
            &bytes,
            Some("legacy.xls".to_string()),
            &ParserConfig::default().zip_config,
            &ArtifactExtractionOptions::default(),
        );

        assert!(bundle.manifest.artifacts.iter().any(|artifact| {
            artifact.kind == ExtractedArtifactKind::OleObject
                && artifact.source_path.as_deref() == Some("Package")
        }));
        assert!(bundle.manifest.artifacts.iter().any(|artifact| {
            artifact.kind == ExtractedArtifactKind::OleNativePayload
                && artifact.mime_type.as_deref() == Some("application/pdf")
        }));
    }

    #[test]
    fn extract_artifacts_finds_legacy_ppt_package_payload() {
        let bytes = build_test_cfb(&[
            ("PowerPoint Document", b"ppt"),
            ("Current User", b"user"),
            ("Package", b"%PDF-1.7"),
        ]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app.parse_bytes(&bytes).expect("parse legacy ppt");
        let bundle = extract_artifacts_from_bytes(
            &parsed,
            &bytes,
            Some("legacy.ppt".to_string()),
            &ParserConfig::default().zip_config,
            &ArtifactExtractionOptions::default(),
        );

        assert!(bundle.manifest.artifacts.iter().any(|artifact| {
            artifact.kind == ExtractedArtifactKind::OleObject
                && artifact.source_path.as_deref() == Some("Package")
        }));
        assert!(bundle.manifest.artifacts.iter().any(|artifact| {
            artifact.kind == ExtractedArtifactKind::OleNativePayload
                && artifact.mime_type.as_deref() == Some("application/pdf")
        }));
    }

    #[test]
    fn extract_artifacts_finds_objectpool_package_payload() {
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            ("ObjectPool/1/Package", b"%PDF-1.7"),
        ]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app.parse_bytes(&bytes).expect("parse objectpool package");
        let bundle = extract_artifacts_from_bytes(
            &parsed,
            &bytes,
            Some("objectpool.doc".to_string()),
            &ParserConfig::default().zip_config,
            &ArtifactExtractionOptions::default(),
        );

        assert!(bundle.manifest.artifacts.iter().any(|artifact| {
            artifact.kind == ExtractedArtifactKind::OleObject
                && artifact.source_path.as_deref() == Some("ObjectPool/1/Package")
        }));
        assert!(bundle.manifest.artifacts.iter().any(|artifact| {
            artifact.kind == ExtractedArtifactKind::OleNativePayload
                && artifact.source_path.as_deref()
                    == Some("ObjectPool/1/Package#ObjectPool/1/Package")
        }));
    }

    #[test]
    fn extract_artifacts_finds_objectpool_ole10native_payload() {
        let mut ole10 = Vec::new();
        ole10.extend_from_slice(&64u32.to_le_bytes());
        ole10.extend_from_slice(&2u16.to_le_bytes());
        ole10.extend_from_slice(b"dropper.exe\0");
        ole10.extend_from_slice(b"C:\\src\\dropper.exe\0");
        ole10.extend_from_slice(&0u32.to_le_bytes());
        ole10.extend_from_slice(&0u32.to_le_bytes());
        ole10.extend_from_slice(b"C:\\temp\\dropper.exe\0");
        ole10.extend_from_slice(&4u32.to_le_bytes());
        ole10.extend_from_slice(b"MZ!!");

        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            ("ObjectPool/1/Ole10Native", &ole10),
        ]);
        let app = DocirApp::new(ParserConfig::default());
        let parsed = app
            .parse_bytes(&bytes)
            .expect("parse objectpool ole10native");
        let bundle = extract_artifacts_from_bytes(
            &parsed,
            &bytes,
            Some("objectpool.doc".to_string()),
            &ParserConfig::default().zip_config,
            &ArtifactExtractionOptions::default(),
        );

        assert!(bundle.manifest.artifacts.iter().any(|artifact| {
            artifact.kind == ExtractedArtifactKind::OleObject
                && artifact.source_path.as_deref() == Some("ObjectPool/1/Ole10Native")
        }));
        assert!(bundle.manifest.artifacts.iter().any(|artifact| {
            artifact.kind == ExtractedArtifactKind::OleNativePayload
                && artifact.suggested_name.as_deref() == Some("dropper.exe")
        }));
    }
}
