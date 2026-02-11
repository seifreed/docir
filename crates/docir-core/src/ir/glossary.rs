//! Word glossary document (building blocks).

use crate::types::{NodeId, SourceSpan};
use serde::{Deserialize, Serialize};

/// Glossary document containing building blocks.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GlossaryDocument {
    pub id: NodeId,
    pub entries: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl GlossaryDocument {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            entries: Vec::new(),
            span: None,
        }
    }
}

/// Glossary entry (docPart).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GlossaryEntry {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gallery: Option<String>,
    pub content: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl GlossaryEntry {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            gallery: None,
            content: Vec::new(),
            span: None,
        }
    }
}
