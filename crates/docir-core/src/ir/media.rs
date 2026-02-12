//! Media asset nodes.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Media asset in the package.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct MediaAsset {
    /// Unique identifier for this node.
    pub id: NodeId,
    /// Path within the package.
    pub path: String,
    /// Media type.
    pub media_type: MediaType,
    /// Content type if known.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub content_type: Option<String>,
    /// Size in bytes.
    pub size_bytes: u64,
    /// Optional hash (hex).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hash: Option<String>,
    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl MediaAsset {
    pub fn new(path: impl Into<String>, media_type: MediaType, size_bytes: u64) -> Self {
        Self {
            id: NodeId::new(),
            path: path.into(),
            media_type,
            content_type: None,
            size_bytes,
            hash: None,
            span: None,
        }
    }
}

/// Media type classification.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MediaType {
    Image,
    Audio,
    Video,
    Other,
}
