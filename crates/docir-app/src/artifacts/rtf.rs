use docir_core::{ExtractedArtifact, ExtractedArtifactKind};
use docir_parser::ole::is_ole_container;

use super::{ArtifactExtractionBundle, ArtifactExtractionOptions, ExtractedPayload};
use crate::artifacts::classify::classify_payload;
use crate::artifacts::helpers::{assign_sha256, preferred_output_name};
use crate::artifacts::ole::extract_embedded_payload;

pub(super) fn extract_rtf_artifacts(
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

pub(crate) fn scan_rtf_objdata(data: &[u8]) -> Vec<Vec<u8>> {
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

pub(crate) const MAX_HEX_BLOB_SIZE: usize = 100 * 1024 * 1024;

pub(crate) fn decode_hex_blob(hex: &[u8]) -> Option<Vec<u8>> {
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

pub(crate) fn hex_val(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}
