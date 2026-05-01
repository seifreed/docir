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
mod tests {
    use super::*;
    use crate::{
        ArtifactInventory, ContainerKind, InventoryArtifact, InventoryArtifactKind,
        VbaModuleReport, VbaProjectReport, VbaRecognitionStatus,
    };

    #[test]
    fn phase0_vba_export_maps_statuses_and_paths() {
        let report = VbaRecognitionReport {
            document_format: "docm".to_string(),
            status: VbaRecognitionStatus::Extracted,
            projects: vec![VbaProjectReport {
                node_id: "node_1".to_string(),
                project_name: Some("Project".to_string()),
                container_path: Some("word/vbaProject.bin".to_string()),
                storage_root: Some("VBA".to_string()),
                is_protected: false,
                has_auto_exec: false,
                auto_exec_procedures: Vec::new(),
                references: Vec::new(),
                status: VbaRecognitionStatus::Extracted,
                modules: vec![VbaModuleReport {
                    node_id: "node_2".to_string(),
                    name: "Module1".to_string(),
                    kind: "standard".to_string(),
                    stream_name: Some("Module1".to_string()),
                    stream_path: Some("VBA/Module1".to_string()),
                    status: VbaRecognitionStatus::Extracted,
                    source_text: Some("Sub Test()\nEnd Sub\n".to_string()),
                    source_encoding: Some("utf-8".to_string()),
                    source_hash: Some("a".repeat(64)),
                    compressed_size: Some(10),
                    decompressed_size: Some(18),
                    procedures: Vec::new(),
                    suspicious_calls: Vec::new(),
                    extraction_errors: Vec::new(),
                }],
            }],
        };

        let export = Phase0VbaExport::from_report(
            &report,
            ExportDocumentRef::new("doc", "docm", Some("/tmp/doc.docm".to_string())),
        );
        assert_eq!(export.vba.status, "extracted");
        assert_eq!(
            export.vba.projects[0].modules[0].storage_path,
            "VBA/Module1"
        );
        assert!(!export.vba.projects[0].is_protected);
    }

    #[test]
    fn phase0_vba_export_omits_source_text_for_recognition_only() {
        let report = VbaRecognitionReport {
            document_format: "doc".to_string(),
            status: VbaRecognitionStatus::Recognized,
            projects: vec![VbaProjectReport {
                node_id: "node_1".to_string(),
                project_name: Some("RecognizeOnly".to_string()),
                container_path: Some("VBA".to_string()),
                storage_root: Some("VBA".to_string()),
                is_protected: false,
                has_auto_exec: true,
                auto_exec_procedures: vec!["AutoOpen".to_string()],
                references: Vec::new(),
                status: VbaRecognitionStatus::Recognized,
                modules: vec![VbaModuleReport {
                    node_id: "node_2".to_string(),
                    name: "Core".to_string(),
                    kind: "standard".to_string(),
                    stream_name: Some("Core".to_string()),
                    stream_path: Some("VBA/Core".to_string()),
                    status: VbaRecognitionStatus::Recognized,
                    source_text: None,
                    source_encoding: None,
                    source_hash: None,
                    compressed_size: Some(24),
                    decompressed_size: Some(24),
                    procedures: vec!["AutoOpen".to_string()],
                    suspicious_calls: Vec::new(),
                    extraction_errors: Vec::new(),
                }],
            }],
        };

        let export = Phase0VbaExport::from_report(
            &report,
            ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
        );

        assert_eq!(export.vba.status, "recognized");
        assert_eq!(export.vba.projects[0].modules[0].status, "recognized");
        assert!(export.vba.projects[0].modules[0].source_text.is_none());
        assert!(export.vba.projects[0].modules[0].source_encoding.is_none());
        assert!(!export.vba.projects[0].is_protected);
    }

    #[test]
    fn phase0_vba_export_keeps_source_text_when_extracted() {
        let report = VbaRecognitionReport {
            document_format: "doc".to_string(),
            status: VbaRecognitionStatus::Extracted,
            projects: vec![VbaProjectReport {
                node_id: "node_1".to_string(),
                project_name: Some("RecognizeWithSource".to_string()),
                container_path: Some("VBA".to_string()),
                storage_root: Some("VBA".to_string()),
                is_protected: false,
                has_auto_exec: true,
                auto_exec_procedures: vec!["AutoOpen".to_string()],
                references: Vec::new(),
                status: VbaRecognitionStatus::Extracted,
                modules: vec![VbaModuleReport {
                    node_id: "node_2".to_string(),
                    name: "Core".to_string(),
                    kind: "standard".to_string(),
                    stream_name: Some("Core".to_string()),
                    stream_path: Some("VBA/Core".to_string()),
                    status: VbaRecognitionStatus::Extracted,
                    source_text: Some("Sub AutoOpen()\nEnd Sub\n".to_string()),
                    source_encoding: Some("utf-8-lossy-normalized".to_string()),
                    source_hash: Some("deadbeef".to_string()),
                    compressed_size: Some(24),
                    decompressed_size: Some(24),
                    procedures: vec!["AutoOpen".to_string()],
                    suspicious_calls: Vec::new(),
                    extraction_errors: Vec::new(),
                }],
            }],
        };

        let export = Phase0VbaExport::from_report(
            &report,
            ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
        );

        assert_eq!(export.vba.status, "extracted");
        assert_eq!(export.vba.projects[0].modules[0].status, "extracted");
        assert_eq!(
            export.vba.projects[0].modules[0].source_text.as_deref(),
            Some("Sub AutoOpen()\nEnd Sub\n")
        );
        assert_eq!(
            export.vba.projects[0].modules[0].source_encoding.as_deref(),
            Some("compressed-container-decoded")
        );
        assert!(!export.vba.projects[0].is_protected);
    }

    #[test]
    fn phase0_vba_export_keeps_partial_status_without_source_text() {
        let report = VbaRecognitionReport {
            document_format: "doc".to_string(),
            status: VbaRecognitionStatus::Partial,
            projects: vec![VbaProjectReport {
                node_id: "node_1".to_string(),
                project_name: Some("PartialRecognizeOnly".to_string()),
                container_path: Some("VBA".to_string()),
                storage_root: Some("VBA".to_string()),
                is_protected: false,
                has_auto_exec: true,
                auto_exec_procedures: vec!["AutoOpen".to_string()],
                references: Vec::new(),
                status: VbaRecognitionStatus::Partial,
                modules: vec![
                    VbaModuleReport {
                        node_id: "node_2".to_string(),
                        name: "Core".to_string(),
                        kind: "standard".to_string(),
                        stream_name: Some("Core".to_string()),
                        stream_path: Some("VBA/Core".to_string()),
                        status: VbaRecognitionStatus::Recognized,
                        source_text: None,
                        source_encoding: None,
                        source_hash: None,
                        compressed_size: Some(24),
                        decompressed_size: Some(24),
                        procedures: vec!["AutoOpen".to_string()],
                        suspicious_calls: Vec::new(),
                        extraction_errors: Vec::new(),
                    },
                    VbaModuleReport {
                        node_id: "node_3".to_string(),
                        name: "MissingMod".to_string(),
                        kind: "standard".to_string(),
                        stream_name: Some("MissingMod".to_string()),
                        stream_path: Some("VBA/MissingMod".to_string()),
                        status: VbaRecognitionStatus::Partial,
                        source_text: None,
                        source_encoding: None,
                        source_hash: None,
                        compressed_size: None,
                        decompressed_size: None,
                        procedures: Vec::new(),
                        suspicious_calls: Vec::new(),
                        extraction_errors: vec!["Missing stream VBA/MissingMod".to_string()],
                    },
                ],
            }],
        };

        let export = Phase0VbaExport::from_report(
            &report,
            ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
        );

        assert_eq!(export.vba.status, "partial");
        assert_eq!(export.vba.projects[0].status, "partial");
        assert_eq!(export.vba.projects[0].modules[0].status, "recognized");
        assert_eq!(export.vba.projects[0].modules[1].status, "partial");
        assert!(export.vba.projects[0].modules[1].source_text.is_none());
        assert_eq!(export.vba.projects[0].modules[1].diagnostics.len(), 1);
        assert!(!export.vba.projects[0].is_protected);
    }

    #[test]
    fn phase0_vba_export_keeps_error_status_without_source_text() {
        let report = VbaRecognitionReport {
            document_format: "doc".to_string(),
            status: VbaRecognitionStatus::Error,
            projects: vec![VbaProjectReport {
                node_id: "node_1".to_string(),
                project_name: Some("DecodeFailRecognizeOnly".to_string()),
                container_path: Some("VBA".to_string()),
                storage_root: Some("VBA".to_string()),
                is_protected: false,
                has_auto_exec: false,
                auto_exec_procedures: Vec::new(),
                references: Vec::new(),
                status: VbaRecognitionStatus::Error,
                modules: vec![VbaModuleReport {
                    node_id: "node_2".to_string(),
                    name: "Broken".to_string(),
                    kind: "standard".to_string(),
                    stream_name: Some("Broken".to_string()),
                    stream_path: Some("VBA/Broken".to_string()),
                    status: VbaRecognitionStatus::Error,
                    source_text: None,
                    source_encoding: None,
                    source_hash: None,
                    compressed_size: Some(0),
                    decompressed_size: None,
                    procedures: Vec::new(),
                    suspicious_calls: Vec::new(),
                    extraction_errors: vec!["Failed to decompress VBA/Broken".to_string()],
                }],
            }],
        };

        let export = Phase0VbaExport::from_report(
            &report,
            ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
        );

        assert_eq!(export.vba.status, "error");
        assert_eq!(export.vba.projects[0].status, "error");
        assert_eq!(export.vba.projects[0].modules[0].status, "error");
        assert!(export.vba.projects[0].modules[0].source_text.is_none());
        assert_eq!(export.vba.projects[0].modules[0].diagnostics.len(), 1);
        assert_eq!(
            export.vba.projects[0].modules[0].diagnostics[0].message,
            "Failed to decompress VBA/Broken"
        );
        assert!(!export.vba.projects[0].is_protected);
    }

    #[test]
    fn phase0_vba_export_preserves_protection_and_references_for_partial_project() {
        let report = VbaRecognitionReport {
            document_format: "doc".to_string(),
            status: VbaRecognitionStatus::Partial,
            projects: vec![VbaProjectReport {
                node_id: "node_1".to_string(),
                project_name: Some("ProtectedPartial".to_string()),
                container_path: Some("VBA".to_string()),
                storage_root: Some("VBA".to_string()),
                is_protected: true,
                has_auto_exec: true,
                auto_exec_procedures: vec!["AutoOpen".to_string()],
                references: vec![
                    r#"Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation"#.to_string(),
                    r#"Reference=*\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#scrrun.dll#Microsoft Scripting Runtime"#.to_string(),
                ],
                status: VbaRecognitionStatus::Partial,
                modules: vec![VbaModuleReport {
                    node_id: "node_2".to_string(),
                    name: "Core".to_string(),
                    kind: "standard".to_string(),
                    stream_name: Some("Core".to_string()),
                    stream_path: Some("VBA/Core".to_string()),
                    status: VbaRecognitionStatus::Partial,
                    source_text: None,
                    source_encoding: None,
                    source_hash: None,
                    compressed_size: Some(24),
                    decompressed_size: None,
                    procedures: vec!["AutoOpen".to_string()],
                    suspicious_calls: Vec::new(),
                    extraction_errors: vec!["Missing stream VBA/Core".to_string()],
                }],
            }],
        };

        let export = Phase0VbaExport::from_report(
            &report,
            ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
        );

        assert!(export.vba.projects[0].is_protected);
        assert_eq!(export.vba.projects[0].references.len(), 2);
        assert!(export.vba.projects[0].references[0].contains("OLE Automation"));
        assert!(export.vba.projects[0].references[1].contains("Microsoft Scripting Runtime"));
    }

    #[test]
    fn phase0_vba_export_preserves_mixed_module_kinds_for_partial_project() {
        let report = VbaRecognitionReport {
            document_format: "doc".to_string(),
            status: VbaRecognitionStatus::Partial,
            projects: vec![VbaProjectReport {
                node_id: "node_1".to_string(),
                project_name: Some("MixedProtectedPartial".to_string()),
                container_path: Some("VBA".to_string()),
                storage_root: Some("VBA".to_string()),
                is_protected: true,
                has_auto_exec: true,
                auto_exec_procedures: vec!["AutoOpen".to_string()],
                references: vec![
                    r#"Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation"#.to_string(),
                ],
                status: VbaRecognitionStatus::Partial,
                modules: vec![
                    VbaModuleReport {
                        node_id: "node_2".to_string(),
                        name: "Core".to_string(),
                        kind: "standard".to_string(),
                        stream_name: Some("Core".to_string()),
                        stream_path: Some("VBA/Core".to_string()),
                        status: VbaRecognitionStatus::Recognized,
                        source_text: None,
                        source_encoding: None,
                        source_hash: None,
                        compressed_size: Some(24),
                        decompressed_size: Some(24),
                        procedures: vec!["AutoOpen".to_string()],
                        suspicious_calls: Vec::new(),
                        extraction_errors: Vec::new(),
                    },
                    VbaModuleReport {
                        node_id: "node_3".to_string(),
                        name: "Helper".to_string(),
                        kind: "class".to_string(),
                        stream_name: Some("Helper".to_string()),
                        stream_path: Some("VBA/Helper".to_string()),
                        status: VbaRecognitionStatus::Partial,
                        source_text: None,
                        source_encoding: None,
                        source_hash: None,
                        compressed_size: None,
                        decompressed_size: None,
                        procedures: Vec::new(),
                        suspicious_calls: Vec::new(),
                        extraction_errors: vec!["Missing stream VBA/Helper".to_string()],
                    },
                    VbaModuleReport {
                        node_id: "node_4".to_string(),
                        name: "ThisDocument".to_string(),
                        kind: "document".to_string(),
                        stream_name: Some("ThisDocument".to_string()),
                        stream_path: Some("VBA/ThisDocument".to_string()),
                        status: VbaRecognitionStatus::Recognized,
                        source_text: None,
                        source_encoding: None,
                        source_hash: None,
                        compressed_size: Some(32),
                        decompressed_size: Some(32),
                        procedures: vec!["Document_Open".to_string()],
                        suspicious_calls: Vec::new(),
                        extraction_errors: Vec::new(),
                    },
                ],
            }],
        };

        let export = Phase0VbaExport::from_report(
            &report,
            ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
        );

        assert_eq!(export.vba.status, "partial");
        assert!(export.vba.projects[0].is_protected);
        assert_eq!(export.vba.projects[0].references.len(), 1);
        assert_eq!(export.vba.projects[0].modules[0].kind, "standard");
        assert_eq!(export.vba.projects[0].modules[1].kind, "class");
        assert_eq!(export.vba.projects[0].modules[1].status, "partial");
        assert_eq!(export.vba.projects[0].modules[2].kind, "document");
        assert_eq!(export.vba.projects[0].modules[2].status, "recognized");
    }

    #[test]
    fn phase0_vba_export_preserves_mixed_module_kinds_for_error_project() {
        let report = VbaRecognitionReport {
            document_format: "doc".to_string(),
            status: VbaRecognitionStatus::Error,
            projects: vec![VbaProjectReport {
                node_id: "node_1".to_string(),
                project_name: Some("MixedProtectedError".to_string()),
                container_path: Some("VBA".to_string()),
                storage_root: Some("VBA".to_string()),
                is_protected: true,
                has_auto_exec: true,
                auto_exec_procedures: vec!["AutoOpen".to_string()],
                references: vec![
                    r#"Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation"#.to_string(),
                ],
                status: VbaRecognitionStatus::Error,
                modules: vec![
                    VbaModuleReport {
                        node_id: "node_2".to_string(),
                        name: "Core".to_string(),
                        kind: "standard".to_string(),
                        stream_name: Some("Core".to_string()),
                        stream_path: Some("VBA/Core".to_string()),
                        status: VbaRecognitionStatus::Recognized,
                        source_text: None,
                        source_encoding: None,
                        source_hash: None,
                        compressed_size: Some(24),
                        decompressed_size: Some(24),
                        procedures: vec!["AutoOpen".to_string()],
                        suspicious_calls: Vec::new(),
                        extraction_errors: Vec::new(),
                    },
                    VbaModuleReport {
                        node_id: "node_3".to_string(),
                        name: "ThisDocument".to_string(),
                        kind: "document".to_string(),
                        stream_name: Some("ThisDocument".to_string()),
                        stream_path: Some("VBA/ThisDocument".to_string()),
                        status: VbaRecognitionStatus::Error,
                        source_text: None,
                        source_encoding: None,
                        source_hash: None,
                        compressed_size: Some(0),
                        decompressed_size: None,
                        procedures: Vec::new(),
                        suspicious_calls: Vec::new(),
                        extraction_errors: vec!["Failed to decompress VBA/ThisDocument".to_string()],
                    },
                ],
            }],
        };

        let export = Phase0VbaExport::from_report(
            &report,
            ExportDocumentRef::new("sample.doc", "doc", Some("sample.doc".to_string())),
        );

        assert_eq!(export.vba.status, "error");
        assert!(export.vba.projects[0].is_protected);
        assert_eq!(export.vba.projects[0].references.len(), 1);
        assert_eq!(export.vba.projects[0].modules[0].kind, "standard");
        assert_eq!(export.vba.projects[0].modules[0].status, "recognized");
        assert_eq!(export.vba.projects[0].modules[1].kind, "document");
        assert_eq!(export.vba.projects[0].modules[1].status, "error");
        assert_eq!(export.vba.projects[0].modules[1].diagnostics.len(), 1);
    }

    #[test]
    fn phase0_manifest_from_inventory_maps_cfb_metadata() {
        let inventory = ArtifactInventory {
            document_format: "doc".to_string(),
            container_kind: ContainerKind::CfbOle,
            artifact_count: 1,
            macro_project_count: 0,
            macro_module_count: 0,
            ole_object_count: 0,
            activex_control_count: 0,
            external_reference_count: 0,
            dde_field_count: 0,
            artifacts: vec![InventoryArtifact {
                kind: InventoryArtifactKind::ContainerEntry,
                node_id: None,
                path: Some("ObjectPool/1/Ole10Native".to_string()),
                relationship_id: None,
                size_bytes: Some(128),
                start_sector: Some(7),
                created_filetime: Some(123),
                modified_filetime: Some(456),
                media_type: None,
                sha256: None,
                details: "ole-native-payload".to_string(),
            }],
        };

        let export = Phase0ArtifactManifestExport::from_inventory(
            &inventory,
            ExportDocumentRef::new("doc", "doc", Some("/tmp/sample.doc".to_string())),
        );

        assert_eq!(export.artifacts[0].kind, "container-entry");
        assert_eq!(export.artifacts[0].locator.offset, Some(7));
        assert_eq!(
            export.artifacts[0].locator.container.as_deref(),
            Some("/tmp/sample.doc")
        );
        assert_eq!(export.artifacts[0].diagnostics.len(), 2);
    }

    #[test]
    fn phase0_manifest_from_inventory_preserves_vba_role_details() {
        let inventory = ArtifactInventory {
            document_format: "doc".to_string(),
            container_kind: ContainerKind::CfbOle,
            artifact_count: 2,
            macro_project_count: 1,
            macro_module_count: 1,
            ole_object_count: 0,
            activex_control_count: 0,
            external_reference_count: 0,
            dde_field_count: 0,
            artifacts: vec![
                InventoryArtifact {
                    kind: InventoryArtifactKind::MacroProject,
                    node_id: Some("node_1".to_string()),
                    path: Some("VBA/PROJECT".to_string()),
                    relationship_id: None,
                    size_bytes: Some(2),
                    start_sector: None,
                    created_filetime: None,
                    modified_filetime: None,
                    media_type: None,
                    sha256: None,
                    details:
                        "name=LegacyProject; protected=true; modules=2; references=1; autoexec=AutoOpen"
                            .to_string(),
                },
                InventoryArtifact {
                    kind: InventoryArtifactKind::MacroModule,
                    node_id: Some("node_2".to_string()),
                    path: Some("VBA/Module1".to_string()),
                    relationship_id: None,
                    size_bytes: Some(24),
                    start_sector: None,
                    created_filetime: None,
                    modified_filetime: None,
                    media_type: None,
                    sha256: None,
                    details:
                        "type=standard; state=extracted; procedures=AutoOpen; errors=Missing stream VBA/MissingMod"
                            .to_string(),
                },
            ],
        };

        let export = Phase0ArtifactManifestExport::from_inventory(
            &inventory,
            ExportDocumentRef::new("doc", "doc", Some("/tmp/sample.doc".to_string())),
        );

        assert_eq!(export.artifacts[0].kind, "vba-project");
        assert_eq!(
            export.artifacts[0].role.as_deref(),
            Some("name=LegacyProject; protected=true; modules=2; references=1; autoexec=AutoOpen")
        );
        assert_eq!(export.artifacts[1].kind, "vba-module-source");
        assert_eq!(
            export.artifacts[1].role.as_deref(),
            Some(
                "type=standard; state=extracted; procedures=AutoOpen; errors=Missing stream VBA/MissingMod"
            )
        );
    }

    #[test]
    fn phase0_manifest_from_inventory_preserves_media_type() {
        let inventory = ArtifactInventory {
            document_format: "docx".to_string(),
            container_kind: ContainerKind::ZipOoxml,
            artifact_count: 1,
            macro_project_count: 0,
            macro_module_count: 0,
            ole_object_count: 0,
            activex_control_count: 0,
            external_reference_count: 0,
            dde_field_count: 0,
            artifacts: vec![InventoryArtifact {
                kind: InventoryArtifactKind::MediaAsset,
                node_id: Some("node_1".to_string()),
                path: Some("word/media/image1.png".to_string()),
                relationship_id: None,
                size_bytes: Some(128),
                start_sector: None,
                created_filetime: None,
                modified_filetime: None,
                media_type: Some("image/png".to_string()),
                sha256: None,
                details: "image/png".to_string(),
            }],
        };

        let export = Phase0ArtifactManifestExport::from_inventory(
            &inventory,
            ExportDocumentRef::new("docx", "docx", Some("/tmp/sample.docx".to_string())),
        );

        assert_eq!(export.artifacts[0].kind, "embedded-file");
        assert_eq!(export.artifacts[0].media_type.as_deref(), Some("image/png"));
    }

    #[test]
    fn phase0_manifest_from_inventory_preserves_sha256() {
        let inventory = ArtifactInventory {
            document_format: "doc".to_string(),
            container_kind: ContainerKind::CfbOle,
            artifact_count: 2,
            macro_project_count: 0,
            macro_module_count: 1,
            ole_object_count: 1,
            activex_control_count: 0,
            external_reference_count: 0,
            dde_field_count: 0,
            artifacts: vec![
                InventoryArtifact {
                    kind: InventoryArtifactKind::MacroModule,
                    node_id: Some("node_1".to_string()),
                    path: Some("VBA/Module1".to_string()),
                    relationship_id: None,
                    size_bytes: Some(24),
                    start_sector: None,
                    created_filetime: None,
                    modified_filetime: None,
                    media_type: None,
                    sha256: Some("a".repeat(64)),
                    details: "type=standard; state=extracted".to_string(),
                },
                InventoryArtifact {
                    kind: InventoryArtifactKind::OleObject,
                    node_id: Some("node_2".to_string()),
                    path: Some("ObjectPool/1/Ole10Native".to_string()),
                    relationship_id: None,
                    size_bytes: Some(128),
                    start_sector: Some(7),
                    created_filetime: None,
                    modified_filetime: None,
                    media_type: Some("application/x-dosexec".to_string()),
                    sha256: Some("b".repeat(64)),
                    details: "dropper.exe".to_string(),
                },
            ],
        };

        let export = Phase0ArtifactManifestExport::from_inventory(
            &inventory,
            ExportDocumentRef::new("doc", "doc", Some("/tmp/sample.doc".to_string())),
        );

        let hash_a = "a".repeat(64);
        let hash_b = "b".repeat(64);
        assert_eq!(export.artifacts[0].sha256.as_deref(), Some(hash_a.as_str()));
        assert_eq!(export.artifacts[1].sha256.as_deref(), Some(hash_b.as_str()));
    }

    #[test]
    fn phase0_manifest_from_extraction_maps_legacy_cfb_metadata() {
        let mut manifest = ExtractionManifest::new();
        manifest.source_document = Some("/tmp/sample.doc".to_string());
        let mut artifact = docir_core::ExtractedArtifact::new(
            "legacy-ole-object-1",
            docir_core::ExtractedArtifactKind::OleObject,
        );
        artifact.source_path = Some("ObjectPool/1/Ole10Native".to_string());
        artifact.size_bytes = Some(64);
        artifact.start_sector = Some(9);
        artifact.created_filetime = Some(111);
        artifact.modified_filetime = Some(222);
        manifest.artifacts.push(artifact);

        let export = Phase0ArtifactManifestExport::from_manifest(
            &manifest,
            ExportDocumentRef::new("doc", "doc", Some("/tmp/sample.doc".to_string())),
        );

        assert_eq!(export.artifacts[0].locator.offset, Some(9));
        assert_eq!(export.artifacts[0].diagnostics.len(), 2);
    }
}
