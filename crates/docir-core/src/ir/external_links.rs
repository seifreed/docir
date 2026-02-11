//! Spreadsheet external links.

use crate::types::{NodeId, SourceSpan};
use serde::{Deserialize, Serialize};

/// External link part (xl/externalLinks/externalLink*.xml).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalLinkPart {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub link_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub target: Option<String>,
    pub sheets: Vec<ExternalLinkSheet>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl ExternalLinkPart {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            link_type: None,
            target: None,
            sheets: Vec::new(),
            span: None,
        }
    }
}

/// External link sheet info.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalLinkSheet {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub r_id: Option<String>,
}
