use crate::{ArtifactInventory, InventoryArtifactKind, VbaRecognitionReport};
use docir_core::ExtractionManifest;
use serde::Serialize;

/// Stable Phase 0 capability flags shared by JSON exports.
#[derive(Debug, Clone, Serialize)]
pub struct PhaseCapabilities {
    pub parsing: bool,
    pub extraction: bool,
    pub vba_recognition: bool,
    pub ast: bool,
    pub deobfuscation: bool,
    pub emulation: bool,
}

impl PhaseCapabilities {
    /// Returns the canonical Phase 0 capability set.
    pub fn phase_0() -> Self {
        Self {
            parsing: true,
            extraction: true,
            vba_recognition: true,
            ast: false,
            deobfuscation: false,
            emulation: false,
        }
    }
}

/// Stable document reference used by schema exports.
#[derive(Debug, Clone, Serialize)]
pub struct ExportDocumentRef {
    pub id: String,
    pub format: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sha256: Option<String>,
}

impl ExportDocumentRef {
    /// Builds a document reference from an input path and normalized format label.
    pub fn new(
        id: impl Into<String>,
        format: impl Into<String>,
        source_path: Option<String>,
    ) -> Self {
        Self {
            id: id.into(),
            format: format.into(),
            source_path,
            sha256: None,
        }
    }
}

/// Schema-aligned Phase 0 VBA export.
#[derive(Debug, Clone, Serialize)]
pub struct Phase0VbaExport {
    pub schema_version: &'static str,
    pub phase: &'static str,
    pub document: ExportDocumentRef,
    pub capabilities: PhaseCapabilities,
    pub vba: Phase0VbaBody,
}

#[derive(Debug, Clone, Serialize)]
pub struct Phase0VbaBody {
    pub status: String,
    pub projects: Vec<Phase0VbaProject>,
}

#[derive(Debug, Clone, Serialize)]
pub struct Phase0VbaProject {
    pub status: String,
    pub container_path: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub project_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub project_id: Option<String>,
    pub is_protected: bool,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub references: Vec<String>,
    pub modules: Vec<Phase0VbaModule>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub diagnostics: Vec<Phase0Diagnostic>,
}

#[derive(Debug, Clone, Serialize)]
pub struct Phase0VbaModule {
    pub name: String,
    pub kind: String,
    pub storage_path: String,
    pub status: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_encoding: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sha256: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub diagnostics: Vec<Phase0Diagnostic>,
}

/// Schema-aligned Phase 0 artifact manifest export.
#[derive(Debug, Clone, Serialize)]
pub struct Phase0ArtifactManifestExport {
    pub schema_version: &'static str,
    pub phase: &'static str,
    pub document: ExportDocumentRef,
    pub capabilities: PhaseCapabilities,
    pub artifacts: Vec<Phase0Artifact>,
}

#[derive(Debug, Clone, Serialize)]
pub struct Phase0Artifact {
    pub id: String,
    pub kind: String,
    pub status: String,
    pub locator: Phase0ArtifactLocator,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub media_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sha256: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub role: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub relationships: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub diagnostics: Vec<Phase0Diagnostic>,
}

#[derive(Debug, Clone, Serialize)]
pub struct Phase0ArtifactLocator {
    pub path: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub container: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub length: Option<u64>,
}

#[derive(Debug, Clone, Serialize)]
pub struct Phase0Diagnostic {
    pub code: String,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
}

impl Phase0VbaExport {
    /// Builds a schema-stable VBA export from the application recognition report.
    pub fn from_report(report: &VbaRecognitionReport, document: ExportDocumentRef) -> Self {
        let projects = report
            .projects
            .iter()
            .map(|project| Phase0VbaProject {
                status: lowercase_debug(project.status),
                container_path: project
                    .container_path
                    .clone()
                    .unwrap_or_else(|| "unknown".to_string()),
                project_name: project.project_name.clone(),
                project_id: Some(project.node_id.clone()),
                is_protected: project.is_protected,
                references: project.references.clone(),
                modules: project
                    .modules
                    .iter()
                    .map(|module| Phase0VbaModule {
                        name: module.name.clone(),
                        kind: module.kind.clone(),
                        storage_path: module
                            .stream_path
                            .clone()
                            .or_else(|| module.stream_name.clone())
                            .unwrap_or_else(|| "unknown".to_string()),
                        status: lowercase_debug(module.status),
                        source_text: module.source_text.clone(),
                        source_encoding: module.source_encoding.as_ref().map(|_| {
                            if module.source_text.is_some() {
                                "compressed-container-decoded".to_string()
                            } else {
                                "unknown".to_string()
                            }
                        }),
                        sha256: module.source_hash.clone(),
                        diagnostics: module
                            .extraction_errors
                            .iter()
                            .map(|message| Phase0Diagnostic {
                                code: "MODULE_EXTRACTION".to_string(),
                                message: message.clone(),
                                path: module.stream_path.clone(),
                            })
                            .collect(),
                    })
                    .collect(),
                diagnostics: Vec::new(),
            })
            .collect();

        Self {
            schema_version: "1.0.0",
            phase: "phase-0",
            document,
            capabilities: PhaseCapabilities::phase_0(),
            vba: Phase0VbaBody {
                status: lowercase_debug(report.status),
                projects,
            },
        }
    }
}

impl Phase0ArtifactManifestExport {
    /// Builds a schema-stable artifact manifest from the extraction manifest.
    pub fn from_manifest(manifest: &ExtractionManifest, document: ExportDocumentRef) -> Self {
        let mut artifacts: Vec<Phase0Artifact> = manifest
            .artifacts
            .iter()
            .map(|artifact| Phase0Artifact {
                id: artifact.id.clone(),
                kind: map_extracted_artifact_kind(artifact.kind).to_string(),
                status: if artifact.output_path.is_some() {
                    "extracted".to_string()
                } else if artifact.errors.is_empty() {
                    "discovered".to_string()
                } else {
                    "partial".to_string()
                },
                locator: Phase0ArtifactLocator {
                    path: artifact
                        .source_path
                        .clone()
                        .or_else(|| artifact.output_path.clone())
                        .unwrap_or_else(|| "unknown".to_string()),
                    container: manifest.source_document.clone(),
                    offset: artifact.start_sector.map(u64::from),
                    length: artifact.size_bytes,
                },
                media_type: artifact.mime_type.clone(),
                size: artifact.size_bytes,
                sha256: artifact.sha256.clone(),
                role: artifact.suggested_name.clone(),
                relationships: Vec::new(),
                diagnostics: extraction_diagnostics(artifact),
            })
            .collect();

        artifacts.extend(manifest.warnings.iter().enumerate().map(|(idx, warning)| {
            Phase0Artifact {
                id: format!("warning-{}", idx + 1),
                kind: "other".to_string(),
                status: "partial".to_string(),
                locator: Phase0ArtifactLocator {
                    path: manifest
                        .source_document
                        .clone()
                        .unwrap_or_else(|| "unknown".to_string()),
                    container: manifest.source_document.clone(),
                    offset: None,
                    length: None,
                },
                media_type: None,
                size: None,
                sha256: None,
                role: Some("warning".to_string()),
                relationships: Vec::new(),
                diagnostics: vec![Phase0Diagnostic {
                    code: warning.code.clone(),
                    message: warning.message.clone(),
                    path: None,
                }],
            }
        }));

        Self {
            schema_version: "1.0.0",
            phase: "phase-0",
            document,
            capabilities: PhaseCapabilities::phase_0(),
            artifacts,
        }
    }

    /// Builds a schema-stable manifest from the inventory command output.
    pub fn from_inventory(inventory: &ArtifactInventory, document: ExportDocumentRef) -> Self {
        let document_container = document
            .source_path
            .clone()
            .or_else(|| Some(document.id.clone()));
        let artifacts = inventory
            .artifacts
            .iter()
            .enumerate()
            .map(|(idx, artifact)| Phase0Artifact {
                id: artifact
                    .node_id
                    .clone()
                    .unwrap_or_else(|| format!("inventory-{}", idx + 1)),
                kind: map_inventory_artifact_kind(artifact.kind).to_string(),
                status: "discovered".to_string(),
                locator: Phase0ArtifactLocator {
                    path: artifact
                        .path
                        .clone()
                        .unwrap_or_else(|| "unknown".to_string()),
                    container: document_container.clone(),
                    offset: artifact.start_sector.map(u64::from),
                    length: artifact.size_bytes,
                },
                media_type: artifact.media_type.clone(),
                size: artifact.size_bytes,
                sha256: artifact.sha256.clone(),
                role: Some(artifact.details.clone()),
                relationships: artifact.relationship_id.clone().into_iter().collect(),
                diagnostics: build_inventory_diagnostics(artifact),
            })
            .collect();

        Self {
            schema_version: "1.0.0",
            phase: "phase-0",
            document,
            capabilities: PhaseCapabilities::phase_0(),
            artifacts,
        }
    }
}

fn map_inventory_artifact_kind(kind: InventoryArtifactKind) -> &'static str {
    match kind {
        InventoryArtifactKind::ContainerEntry => "container-entry",
        InventoryArtifactKind::RelationshipGraph => "relationship-target",
        InventoryArtifactKind::ExtensionPart => "container-entry",
        InventoryArtifactKind::MediaAsset => "embedded-file",
        InventoryArtifactKind::CustomXmlPart => "container-entry",
        InventoryArtifactKind::MacroProject => "vba-project",
        InventoryArtifactKind::MacroModule => "vba-module-source",
        InventoryArtifactKind::OleObject => "ole-object",
        InventoryArtifactKind::ActiveXControl => "active-content",
        InventoryArtifactKind::ExternalReference => "relationship-target",
        InventoryArtifactKind::DdeField => "active-content",
        InventoryArtifactKind::Diagnostics => "other",
    }
}

fn build_inventory_diagnostics(artifact: &crate::InventoryArtifact) -> Vec<Phase0Diagnostic> {
    let mut diagnostics = Vec::new();
    if let Some(created) = artifact.created_filetime {
        diagnostics.push(Phase0Diagnostic {
            code: "CFB_CREATED_FILETIME".to_string(),
            message: created.to_string(),
            path: artifact.path.clone(),
        });
    }
    if let Some(modified) = artifact.modified_filetime {
        diagnostics.push(Phase0Diagnostic {
            code: "CFB_MODIFIED_FILETIME".to_string(),
            message: modified.to_string(),
            path: artifact.path.clone(),
        });
    }
    diagnostics
}

fn map_extracted_artifact_kind(kind: docir_core::ExtractedArtifactKind) -> &'static str {
    match kind {
        docir_core::ExtractedArtifactKind::MacroProject => "vba-project",
        docir_core::ExtractedArtifactKind::MacroModule => "vba-module-source",
        docir_core::ExtractedArtifactKind::MediaAsset => "embedded-file",
        docir_core::ExtractedArtifactKind::OleObject => "ole-object",
        docir_core::ExtractedArtifactKind::OleNativePayload => "embedded-file",
        docir_core::ExtractedArtifactKind::ActiveXControl => "active-content",
        docir_core::ExtractedArtifactKind::ExternalReference => "relationship-target",
        docir_core::ExtractedArtifactKind::DdeField => "active-content",
        docir_core::ExtractedArtifactKind::RtfEmbeddedObject => "embedded-file",
    }
}

fn extraction_diagnostics(artifact: &docir_core::ExtractedArtifact) -> Vec<Phase0Diagnostic> {
    let mut diagnostics: Vec<Phase0Diagnostic> = artifact
        .errors
        .iter()
        .map(|message| Phase0Diagnostic {
            code: "EXTRACTION".to_string(),
            message: message.clone(),
            path: artifact.source_path.clone(),
        })
        .collect();

    if let Some(created) = artifact.created_filetime {
        diagnostics.push(Phase0Diagnostic {
            code: "CFB_CREATED_FILETIME".to_string(),
            message: created.to_string(),
            path: artifact.source_path.clone(),
        });
    }
    if let Some(modified) = artifact.modified_filetime {
        diagnostics.push(Phase0Diagnostic {
            code: "CFB_MODIFIED_FILETIME".to_string(),
            message: modified.to_string(),
            path: artifact.source_path.clone(),
        });
    }
    diagnostics
}

fn lowercase_debug<T: std::fmt::Debug>(value: T) -> String {
    format!("{:?}", value).to_ascii_lowercase()
}

#[cfg(test)]
mod tests;
