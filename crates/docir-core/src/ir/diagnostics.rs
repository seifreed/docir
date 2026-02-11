//! Diagnostics IR nodes.

use crate::types::{NodeId, SourceSpan};
use serde::{Deserialize, Serialize};

/// Diagnostic severity levels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum DiagnosticSeverity {
    Info,
    Warning,
    Error,
}

/// Diagnostic entry.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiagnosticEntry {
    pub severity: DiagnosticSeverity,
    pub code: String,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
}

/// Diagnostics container.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Diagnostics {
    pub id: NodeId,
    pub entries: Vec<DiagnosticEntry>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Diagnostics {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            entries: Vec::new(),
            span: None,
        }
    }
}
