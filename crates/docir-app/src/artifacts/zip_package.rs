use docir_core::{ExtractedArtifact, ExtractedArtifactKind, ExtractionWarning};
use docir_parser::zip_handler::{SecureZipReader, ZipConfig};
use std::collections::HashSet;
use std::io::Cursor;

use super::{ArtifactExtractionBundle, ArtifactExtractionOptions, ExtractedPayload};
use crate::artifacts::classify::{classify_media_asset, classify_payload};
use crate::artifacts::helpers::{
    assign_sha256, file_name_from_path, preferred_output_name, sanitize_name,
};
use crate::artifacts::ole::{EmbeddedPayload, extract_embedded_payload};

struct ZipArtifactCandidate {
    path: String,
    kind: ExtractedArtifactKind,
    prefix: &'static str,
}

pub(super) fn extract_odf_artifacts(
    input_bytes: &[u8],
    zip_config: &ZipConfig,
    options: &ArtifactExtractionOptions,
    bundle: &mut ArtifactExtractionBundle,
) {
    let Some(mut zip) = open_zip(input_bytes, zip_config, "ODF", bundle) else {
        return;
    };
    let paths = collect_odf_artifact_paths(&zip, options);
    extract_zip_artifacts(&mut zip, paths, options, bundle);
}

pub(super) fn extract_hwpx_artifacts(
    input_bytes: &[u8],
    zip_config: &ZipConfig,
    options: &ArtifactExtractionOptions,
    bundle: &mut ArtifactExtractionBundle,
) {
    let Some(mut zip) = open_zip(input_bytes, zip_config, "HWPX", bundle) else {
        return;
    };
    let paths = collect_hwpx_artifact_paths(&zip, options);
    extract_zip_artifacts(&mut zip, paths, options, bundle);
}

fn open_zip<'a>(
    input_bytes: &'a [u8],
    zip_config: &ZipConfig,
    label: &str,
    bundle: &mut ArtifactExtractionBundle,
) -> Option<SecureZipReader<Cursor<&'a [u8]>>> {
    match SecureZipReader::new(Cursor::new(input_bytes), zip_config.clone()) {
        Ok(zip) => Some(zip),
        Err(err) => {
            bundle.manifest.warnings.push(ExtractionWarning::new(
                "ZIP_OPEN_FAILED",
                format!("Unable to reopen the {label} package for artifact extraction: {err}"),
            ));
            None
        }
    }
}

fn collect_odf_artifact_paths<R: std::io::Read + std::io::Seek>(
    zip: &SecureZipReader<R>,
    options: &ArtifactExtractionOptions,
) -> Vec<ZipArtifactCandidate> {
    let mut paths = Vec::new();
    paths.extend(collect_prefixed_paths(
        zip,
        "ObjectReplacements/",
        ExtractedArtifactKind::OleObject,
        "odf-ole-object",
    ));
    paths.extend(collect_object_paths(zip));
    if !options.no_media {
        paths.extend(collect_prefixed_paths(
            zip,
            "Pictures/",
            ExtractedArtifactKind::MediaAsset,
            "odf-media-asset",
        ));
    }
    sort_and_deduplicate(paths)
}

fn collect_hwpx_artifact_paths<R: std::io::Read + std::io::Seek>(
    zip: &SecureZipReader<R>,
    options: &ArtifactExtractionOptions,
) -> Vec<ZipArtifactCandidate> {
    if options.no_media || options.only_ole {
        return Vec::new();
    }
    sort_and_deduplicate(collect_prefixed_paths(
        zip,
        "BinData/",
        ExtractedArtifactKind::MediaAsset,
        "hwpx-media-asset",
    ))
}

fn collect_object_paths<R: std::io::Read + std::io::Seek>(
    zip: &SecureZipReader<R>,
) -> Vec<ZipArtifactCandidate> {
    zip.file_names()
        .filter(|path| path.starts_with("Object ") && !path.ends_with('/'))
        .map(|path| ZipArtifactCandidate {
            path: path.to_string(),
            kind: ExtractedArtifactKind::OleObject,
            prefix: "odf-ole-object",
        })
        .collect()
}

fn collect_prefixed_paths<R: std::io::Read + std::io::Seek>(
    zip: &SecureZipReader<R>,
    prefix: &str,
    kind: ExtractedArtifactKind,
    id_prefix: &'static str,
) -> Vec<ZipArtifactCandidate> {
    zip.list_prefix(prefix)
        .into_iter()
        .filter(|path| !path.ends_with('/'))
        .map(|path| ZipArtifactCandidate {
            path: path.to_string(),
            kind,
            prefix: id_prefix,
        })
        .collect()
}

fn sort_and_deduplicate(mut paths: Vec<ZipArtifactCandidate>) -> Vec<ZipArtifactCandidate> {
    paths.sort_by(|left, right| left.path.cmp(&right.path));
    let mut seen = HashSet::new();
    paths
        .into_iter()
        .filter(|item| seen.insert(item.path.clone()))
        .collect()
}

fn extract_zip_artifacts<R: std::io::Read + std::io::Seek>(
    zip: &mut SecureZipReader<R>,
    paths: Vec<ZipArtifactCandidate>,
    options: &ArtifactExtractionOptions,
    bundle: &mut ArtifactExtractionBundle,
) {
    let mut ordinal = 0usize;
    let mut payload_index = 0usize;
    for candidate in paths {
        if options.only_ole && candidate.kind != ExtractedArtifactKind::OleObject {
            continue;
        }
        ordinal += 1;
        let data = match zip.read_file(&candidate.path) {
            Ok(data) => data,
            Err(err) => {
                bundle.manifest.warnings.push(ExtractionWarning::new(
                    "ARTIFACT_READ_FAILED",
                    format!(
                        "Unable to read embedded artifact {}: {}",
                        candidate.path, err
                    ),
                ));
                continue;
            }
        };
        let mut artifact = build_zip_artifact(&candidate, ordinal, &data, options);
        if candidate.kind == ExtractedArtifactKind::MediaAsset {
            push_media_payload(&mut artifact, ordinal, data, bundle);
            bundle.manifest.artifacts.push(artifact);
            continue;
        } else if options.with_raw {
            push_raw_payload(&mut artifact, &candidate.path, data.clone(), bundle);
        }
        bundle.manifest.artifacts.push(artifact);

        match extract_embedded_payload(&data) {
            Ok(Some(payload)) => {
                payload_index += 1;
                push_embedded_payload(&candidate.path, payload, payload_index, options, bundle);
            }
            Ok(None) => {}
            Err(err) => bundle.manifest.warnings.push(ExtractionWarning::new(
                "OLE_EMBEDDED_OPEN_FAILED",
                format!(
                    "Unable to open embedded OLE artifact {}: {}",
                    candidate.path, err
                ),
            )),
        }
    }
}

fn build_zip_artifact(
    candidate: &ZipArtifactCandidate,
    ordinal: usize,
    data: &[u8],
    options: &ArtifactExtractionOptions,
) -> ExtractedArtifact {
    let mut artifact =
        ExtractedArtifact::new(format!("{}-{}", candidate.prefix, ordinal), candidate.kind);
    artifact.source_path = Some(candidate.path.clone());
    artifact.suggested_name = Some(file_name_from_path(&candidate.path));
    artifact.size_bytes = Some(data.len() as u64);
    assign_sha256(&mut artifact.sha256, data, options.compute_hashes);
    let (_, mime_type) = if candidate.kind == ExtractedArtifactKind::MediaAsset {
        classify_media_asset(&candidate.path, data)
    } else {
        classify_payload(data, artifact.suggested_name.as_deref())
    };
    artifact.mime_type = Some(mime_type.to_string());
    artifact
}

fn push_media_payload(
    artifact: &mut ExtractedArtifact,
    ordinal: usize,
    data: Vec<u8>,
    bundle: &mut ArtifactExtractionBundle,
) {
    let file_name = artifact
        .suggested_name
        .clone()
        .unwrap_or_else(|| format!("artifact_{ordinal}"));
    let relative_path = format!("payloads/{}", sanitize_name(&file_name));
    artifact.output_path = Some(relative_path.clone());
    bundle.payloads.push(ExtractedPayload {
        artifact_id: artifact.id.clone(),
        relative_path,
        data,
    });
}

fn push_raw_payload(
    artifact: &mut ExtractedArtifact,
    path: &str,
    data: Vec<u8>,
    bundle: &mut ArtifactExtractionBundle,
) {
    let relative_path = format!("raw/zip_{}", sanitize_name(path));
    artifact.output_path = Some(relative_path.clone());
    bundle.payloads.push(ExtractedPayload {
        artifact_id: artifact.id.clone(),
        relative_path,
        data,
    });
}

fn push_embedded_payload(
    path: &str,
    payload: EmbeddedPayload,
    payload_index: usize,
    options: &ArtifactExtractionOptions,
    bundle: &mut ArtifactExtractionBundle,
) {
    let mut artifact = ExtractedArtifact::new(
        format!("zip-ole-native-payload-{}", payload_index),
        ExtractedArtifactKind::OleNativePayload,
    );
    artifact.source_path = Some(format!("{}#{}", path, payload.stream_name));
    artifact.suggested_name = payload.file_name.clone();
    artifact.size_bytes = Some(payload.data.len() as u64);
    assign_sha256(&mut artifact.sha256, &payload.data, options.compute_hashes);
    let (payload_kind, mime_type) = classify_payload(&payload.data, payload.file_name.as_deref());
    artifact.mime_type = Some(mime_type.to_string());
    let file_name = preferred_output_name(
        payload.file_name.as_deref(),
        payload_index,
        payload_kind,
        artifact.mime_type.as_deref(),
    );
    let relative_path = format!("payloads/{}", file_name);
    artifact.output_path = Some(relative_path.clone());
    bundle.payloads.push(ExtractedPayload {
        artifact_id: artifact.id.clone(),
        relative_path,
        data: payload.data,
    });
    bundle.manifest.artifacts.push(artifact);
}
