//! Custom XML parts.

use crate::types::{NodeId, SourceSpan};
use serde::{Deserialize, Serialize};

/// Custom XML part.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomXmlPart {
    /// Unique identifier for this node.
    pub id: NodeId,
    /// Path within the package.
    pub path: String,
    /// Root element name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub root_element: Option<String>,
    /// Namespace URIs detected.
    pub namespaces: Vec<String>,
    /// Size in bytes.
    pub size_bytes: u64,
    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl CustomXmlPart {
    pub fn new(path: impl Into<String>, size_bytes: u64) -> Self {
        Self {
            id: NodeId::new(),
            path: path.into(),
            root_element: None,
            namespaces: Vec::new(),
            size_bytes,
            span: None,
        }
    }
}
