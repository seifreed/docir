//! Office add-ins (web extensions).

use crate::types::{NodeId, SourceSpan};
use serde::{Deserialize, Serialize};

/// Web extension definition (word/webExtensions/webExtension*.xml).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WebExtension {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub extension_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub store: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub store_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub store_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub version: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reference_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reference_version: Option<String>,
    pub properties: Vec<WebExtensionProperty>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl WebExtension {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            extension_id: None,
            store: None,
            store_type: None,
            store_id: None,
            version: None,
            reference_id: None,
            reference_version: None,
            properties: Vec::new(),
            span: None,
        }
    }
}

/// Web extension property (name/value).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WebExtensionProperty {
    pub name: String,
    pub value: String,
}

/// Web extension taskpane definition (word/webExtensions/taskpanes.xml).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WebExtensionTaskpane {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub web_extension_ref: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dock_state: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub visibility: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub row: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub column: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl WebExtensionTaskpane {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            web_extension_ref: None,
            dock_state: None,
            visibility: None,
            width: None,
            height: None,
            row: None,
            column: None,
            span: None,
        }
    }
}
