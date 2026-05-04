use docir_core::{ExtractedArtifact, ExtractedArtifactKind, ExtractionWarning};
use docir_parser::ole::{Cfb, CfbEntryMetadata};

use super::{ArtifactExtractionBundle, ArtifactExtractionOptions, ExtractedPayload};
use crate::artifacts::classify::classify_payload;
use crate::artifacts::helpers::{
    assign_sha256, file_name_from_path, preferred_output_name, sanitize_name,
};
use crate::artifacts::ole::{parse_ole10_native, EmbeddedPayload};

pub(super) fn extract_legacy_cfb_artifacts(
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
        let metadata = cfb.entry_metadata(&path);
        let upper = path.to_ascii_uppercase();
        if !is_legacy_ole_stream(&upper) {
            continue;
        }
        if !passes_legacy_only_ole_filter(&upper, options) {
            continue;
        }

        let Some(data) = cfb.read_stream(&path) else {
            continue;
        };

        payload_index += 1;
        let mut artifact =
            build_legacy_ole_artifact(&path, &data, metadata, payload_index, options);

        if options.with_raw {
            push_legacy_raw_payload(&mut artifact, &path, data.clone(), bundle);
        }

        bundle.manifest.artifacts.push(artifact);

        if let Some(payload) = extract_legacy_payload(&upper, &path, data) {
            payload_index += 1;
            push_legacy_embedded_payload(&path, payload, metadata, payload_index, options, bundle);
        }
    }
}

fn is_legacy_ole_stream(upper_path: &str) -> bool {
    upper_path.contains("OBJECTPOOL/")
        || upper_path.ends_with("OLE10NATIVE")
        || upper_path == "PACKAGE"
        || upper_path.ends_with("/PACKAGE")
        || upper_path.ends_with("/CONTENTS")
}

fn passes_legacy_only_ole_filter(upper_path: &str, options: &ArtifactExtractionOptions) -> bool {
    !options.only_ole
        || upper_path.contains("OBJECTPOOL/")
        || upper_path.ends_with("OLE10NATIVE")
        || upper_path == "PACKAGE"
        || upper_path.ends_with("/PACKAGE")
}

fn build_legacy_ole_artifact(
    path: &str,
    data: &[u8],
    metadata: Option<&CfbEntryMetadata>,
    index: usize,
    options: &ArtifactExtractionOptions,
) -> ExtractedArtifact {
    let mut artifact = ExtractedArtifact::new(
        format!("legacy-ole-object-{}", index),
        ExtractedArtifactKind::OleObject,
    );
    artifact.source_path = Some(path.to_string());
    artifact.suggested_name = Some(file_name_from_path(path));
    artifact.size_bytes = Some(data.len() as u64);
    apply_cfb_metadata(&mut artifact, metadata);
    assign_sha256(&mut artifact.sha256, data, options.compute_hashes);
    let (_, mime_type) = classify_payload(data, artifact.suggested_name.as_deref());
    artifact.mime_type = Some(mime_type.to_string());
    artifact
}

fn push_legacy_raw_payload(
    artifact: &mut ExtractedArtifact,
    path: &str,
    data: Vec<u8>,
    bundle: &mut ArtifactExtractionBundle,
) {
    let raw_name = format!("legacy_{}", sanitize_name(path));
    let relative_path = format!("raw/{}", raw_name);
    artifact.output_path = Some(relative_path.clone());
    bundle.payloads.push(ExtractedPayload {
        artifact_id: artifact.id.clone(),
        relative_path,
        data,
    });
}

fn extract_legacy_payload(upper_path: &str, path: &str, data: Vec<u8>) -> Option<EmbeddedPayload> {
    if upper_path.ends_with("OLE10NATIVE") {
        return parse_ole10_native(&data).map(|payload| EmbeddedPayload {
            stream_name: path.to_string(),
            file_name: payload.file_name,
            source_path: payload.source_path,
            temp_path: payload.temp_path,
            data: payload.data,
        });
    }
    if upper_path.ends_with("/PACKAGE") || upper_path == "PACKAGE" {
        return Some(EmbeddedPayload {
            stream_name: path.to_string(),
            file_name: None,
            source_path: None,
            temp_path: None,
            data,
        });
    }
    crate::artifacts::ole::extract_embedded_payload_from_cfb(&data)
}

fn push_legacy_embedded_payload(
    path: &str,
    payload: EmbeddedPayload,
    metadata: Option<&CfbEntryMetadata>,
    payload_index: usize,
    options: &ArtifactExtractionOptions,
    bundle: &mut ArtifactExtractionBundle,
) {
    let mut payload_artifact = ExtractedArtifact::new(
        format!("legacy-payload-{}", payload_index),
        ExtractedArtifactKind::OleNativePayload,
    );
    payload_artifact.source_path = Some(format!("{}#{}", path, payload.stream_name));
    payload_artifact.suggested_name = payload.file_name.clone();
    payload_artifact.size_bytes = Some(payload.data.len() as u64);
    apply_cfb_metadata(&mut payload_artifact, metadata);
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

fn apply_cfb_metadata(artifact: &mut ExtractedArtifact, metadata: Option<&CfbEntryMetadata>) {
    artifact.start_sector = metadata.map(|entry| entry.start_sector);
    artifact.created_filetime = metadata.and_then(|entry| entry.created_filetime);
    artifact.modified_filetime = metadata.and_then(|entry| entry.modified_filetime);
}
