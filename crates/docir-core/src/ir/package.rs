//! Package-level shared nodes (relationships, extensions).

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Relationship graph for a part.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct RelationshipGraph {
    /// Unique identifier for this node.
    pub id: NodeId,
    /// Source part path (e.g., "_rels/.rels" or "word/document.xml").
    pub source: String,
    /// Relationships from this source.
    pub relationships: Vec<RelationshipEntry>,
    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl RelationshipGraph {
    pub fn new(source: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            source: source.into(),
            relationships: Vec::new(),
            span: None,
        }
    }
}

/// A relationship entry.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct RelationshipEntry {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: RelationshipTargetMode,
}

/// Relationship target mode.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RelationshipTargetMode {
    Internal,
    External,
}

/// Extension part for unknown or legacy content.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ExtensionPart {
    /// Unique identifier for this node.
    pub id: NodeId,
    /// Path within the package.
    pub path: String,
    /// Content type if known.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub content_type: Option<String>,
    /// Size in bytes.
    pub size_bytes: u64,
    /// Kind of extension.
    pub kind: ExtensionPartKind,
    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl ExtensionPart {
    pub fn new(path: impl Into<String>, size_bytes: u64, kind: ExtensionPartKind) -> Self {
        Self {
            id: NodeId::new(),
            path: path.into(),
            content_type: None,
            size_bytes,
            kind,
            span: None,
        }
    }
}

/// Extension part kind.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExtensionPartKind {
    Legacy,
    VendorSpecific,
    Unknown,
}
