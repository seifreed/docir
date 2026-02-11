//! Spreadsheet query tables.

use crate::types::{NodeId, SourceSpan};
use serde::{Deserialize, Serialize};

/// Query table part (xl/queryTables/queryTable*.xml).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct QueryTablePart {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub connection_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub command: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub url: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl QueryTablePart {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            connection_id: None,
            command: None,
            url: None,
            source: None,
            span: None,
        }
    }
}
