//! WordprocessingML settings, web settings, and font table.

use crate::types::{NodeId, SourceSpan};
use serde::{Deserialize, Serialize};

/// Generic settings container for word/settings.xml.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WordSettings {
    pub id: NodeId,
    pub entries: Vec<SettingEntry>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl WordSettings {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            entries: Vec::new(),
            span: None,
        }
    }
}

/// Settings entry (element name + optional value + attributes).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SettingEntry {
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
    pub attributes: Vec<SettingAttribute>,
}

/// Attribute for settings entries.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SettingAttribute {
    pub name: String,
    pub value: String,
}

/// Web settings (word/webSettings.xml).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WebSettings {
    pub id: NodeId,
    pub entries: Vec<SettingEntry>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl WebSettings {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            entries: Vec::new(),
            span: None,
        }
    }
}

/// Font table (word/fontTable.xml).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontTable {
    pub id: NodeId,
    pub fonts: Vec<FontEntry>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl FontTable {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            fonts: Vec::new(),
            span: None,
        }
    }
}

/// Font entry in fontTable.xml.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontEntry {
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alt_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub charset: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub panose: Option<String>,
}
