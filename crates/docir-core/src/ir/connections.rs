//! Spreadsheet connections part (xl/connections.xml).

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Connection part containing all workbook connections.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ConnectionPart {
    pub id: NodeId,
    pub entries: Vec<ConnectionEntry>,
    pub span: Option<SourceSpan>,
}

impl ConnectionPart {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            entries: Vec::new(),
            span: None,
        }
    }
}

/// A single connection entry (from connections.xml).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ConnectionEntry {
    pub connection_id: Option<u32>,
    pub name: Option<String>,
    pub description: Option<String>,
    pub connection_type: Option<u32>,
    pub refreshed_version: Option<u32>,
    pub refresh_on_load: Option<bool>,
    pub save_data: Option<bool>,
    pub background: Option<bool>,
    pub connection: Option<String>,
    pub command: Option<String>,
    pub command_type: Option<u32>,
    pub source_file: Option<String>,
    pub connection_file: Option<String>,
    pub url: Option<String>,
}

impl ConnectionEntry {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            connection_id: None,
            name: None,
            description: None,
            connection_type: None,
            refreshed_version: None,
            refresh_on_load: None,
            save_data: None,
            background: None,
            connection: None,
            command: None,
            command_type: None,
            source_file: None,
            connection_file: None,
            url: None,
        }
    }
}
