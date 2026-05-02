use docir_core::{ExtractedArtifact, ExtractedArtifactKind, ExtractionWarning};
use docir_parser::ole::Cfb;

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
            crate::artifacts::ole::extract_embedded_payload_from_cfb(&data)
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
