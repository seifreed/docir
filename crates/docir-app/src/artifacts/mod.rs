use crate::ParsedDocument;
use docir_core::types::DocumentFormat;
use docir_core::{ExtractionManifest, ExtractionWarning};
use docir_parser::zip_handler::ZipConfig;

mod classify;
mod helpers;
mod legacy_cfb;
mod ole;
mod ooxml;
mod rtf;

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
                legacy_cfb::extract_legacy_cfb_artifacts(input_bytes, options, &mut bundle);
                return bundle;
            }
            if options.only_rtf_objects {
                bundle.manifest.warnings.push(ExtractionWarning::new(
                    "NO_MATCHING_ARTIFACTS",
                    "RTF-only extraction requested for a non-RTF container",
                ));
                return bundle;
            }
            ooxml::extract_ooxml_artifacts(input_bytes, zip_config, options, &mut bundle);
        }
        DocumentFormat::Rtf => {
            rtf::extract_rtf_artifacts(input_bytes, options, &mut bundle);
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

fn is_legacy_cfb_document(parsed: &ParsedDocument) -> bool {
    parsed
        .document()
        .and_then(|doc| doc.span.as_ref())
        .map(|span| span.file_path.starts_with("cfb:/"))
        .unwrap_or(false)
}

pub(crate) use rtf::scan_rtf_objdata;

#[cfg(test)]
mod tests;
