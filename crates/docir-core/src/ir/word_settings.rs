//! WordprocessingML settings, web settings, and font table.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Generic settings container for word/settings.xml.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct WordSettings {
    pub id: NodeId,
    pub entries: Vec<SettingEntry>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl WordSettings {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            entries: Vec::new(),
            span: None,
        }
    }
}

/// Settings entry (element name + optional value + attributes).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SettingEntry {
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
    pub attributes: Vec<SettingAttribute>,
}

/// Attribute for settings entries.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SettingAttribute {
    pub name: String,
    pub value: String,
}

/// Web settings (word/webSettings.xml).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct WebSettings {
    pub id: NodeId,
    pub entries: Vec<SettingEntry>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl WebSettings {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            entries: Vec::new(),
            span: None,
        }
    }
}

/// Font table (word/fontTable.xml).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct FontTable {
    pub id: NodeId,
    pub fonts: Vec<FontEntry>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl FontTable {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            fonts: Vec::new(),
            span: None,
        }
    }
}

/// Font entry in fontTable.xml.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
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
