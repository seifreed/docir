use docir_core::{ExtractedArtifact, ExtractedArtifactKind, ExtractionWarning};
use docir_parser::zip_handler::{SecureZipReader, ZipConfig};
use std::collections::HashSet;
use std::io::Cursor;

use super::{ArtifactExtractionBundle, ArtifactExtractionOptions, ExtractedPayload};
use crate::artifacts::classify::classify_media_asset;
use crate::artifacts::classify::classify_payload;
use crate::artifacts::helpers::{
    assign_sha256, file_name_from_path, preferred_output_name, sanitize_name,
};
use crate::artifacts::ole::extract_embedded_payload;

pub(super) fn extract_ooxml_artifacts(
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
