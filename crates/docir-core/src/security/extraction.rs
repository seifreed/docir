use crate::types::NodeId;
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Serializable extraction manifest shared by CLI, JSON exports, and downstream tooling.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct ExtractionManifest {
    /// Schema identifier for downstream consumers.
    pub schema_version: String,
    /// Source document path or logical identifier.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_document: Option<String>,
    /// Artifacts exported or recognized during the extraction step.
    #[serde(default)]
    pub artifacts: Vec<ExtractedArtifact>,
    /// Non-fatal issues collected while extracting.
    #[serde(default)]
    pub warnings: Vec<ExtractionWarning>,
}

impl ExtractionManifest {
    /// Creates a new manifest with the current schema version.
    pub fn new() -> Self {
        Self {
            schema_version: "docir.extraction.v1".to_string(),
            source_document: None,
            artifacts: Vec::new(),
            warnings: Vec::new(),
        }
    }
}

/// A single extracted or recognized artifact.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ExtractedArtifact {
    /// Stable artifact identifier.
    pub id: String,
    /// Logical artifact kind.
    pub kind: ExtractedArtifactKind,
    /// Related IR node when the artifact originates from a stored node.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub node_id: Option<NodeId>,
    /// Original path within the container or source document.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_path: Option<String>,
    /// Suggested filename for downstream export.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub suggested_name: Option<String>,
    /// Normalized text encoding when the payload is textual.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub encoding: Option<String>,
    /// MIME type estimate for the extracted payload.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mime_type: Option<String>,
    /// SHA-256 hash of the payload when available.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sha256: Option<String>,
    /// Payload size in bytes when known.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size_bytes: Option<u64>,
    /// CFB start sector when the artifact comes from a legacy OLE stream.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start_sector: Option<u32>,
    /// CFB created FILETIME when available.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub created_filetime: Option<u64>,
    /// CFB modified FILETIME when available.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub modified_filetime: Option<u64>,
    /// Relative output path written by the CLI when exported to disk.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub output_path: Option<String>,
    /// Non-fatal extraction issues for this artifact.
    #[serde(default)]
    pub errors: Vec<String>,
}

impl ExtractedArtifact {
    /// Creates a new extracted artifact record.
    pub fn new(id: impl Into<String>, kind: ExtractedArtifactKind) -> Self {
        Self {
            id: id.into(),
            kind,
            node_id: None,
            source_path: None,
            suggested_name: None,
            encoding: None,
            mime_type: None,
            sha256: None,
            size_bytes: None,
            start_sector: None,
            created_filetime: None,
            modified_filetime: None,
            output_path: None,
            errors: Vec::new(),
        }
    }
}

/// High-level artifact taxonomy used by manifests and export commands.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExtractedArtifactKind {
    MacroProject,
    MacroModule,
    MediaAsset,
    OleObject,
    OleNativePayload,
    ActiveXControl,
    ExternalReference,
    DdeField,
    RtfEmbeddedObject,
}

/// Non-fatal extraction warning attached to a manifest.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ExtractionWarning {
    /// Warning code stable enough for downstream automation.
    pub code: String,
    /// Human-readable warning description.
    pub message: String,
}

impl ExtractionWarning {
    /// Creates a new manifest warning.
    pub fn new(code: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            code: code.into(),
            message: message.into(),
        }
    }
}
