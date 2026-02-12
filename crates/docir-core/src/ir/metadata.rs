//! Document metadata IR nodes.

use crate::types::NodeId;
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Document metadata (core properties).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct DocumentMetadata {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Document title.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,

    /// Document subject.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subject: Option<String>,

    /// Document creator/author.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub creator: Option<String>,

    /// Keywords.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub keywords: Option<String>,

    /// Description.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,

    /// Last modified by.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_modified_by: Option<String>,

    /// Revision number.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub revision: Option<String>,

    /// Creation date (ISO 8601).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub created: Option<String>,

    /// Last modified date (ISO 8601).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub modified: Option<String>,

    /// Category.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub category: Option<String>,

    /// Content status.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub content_status: Option<String>,

    /// Application that created the document.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub application: Option<String>,

    /// Application version.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub app_version: Option<String>,

    /// Company name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub company: Option<String>,

    /// Manager.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub manager: Option<String>,

    /// Custom properties.
    pub custom_properties: Vec<CustomProperty>,
}

impl DocumentMetadata {
    /// Creates a new empty DocumentMetadata.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            ..Default::default()
        }
    }
}

/// A custom document property.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct CustomProperty {
    /// Property name.
    pub name: String,

    /// Property value.
    pub value: PropertyValue,

    /// Property format ID (for typed properties).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub format_id: Option<String>,

    /// Property ID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub property_id: Option<u32>,
}

/// Custom property value types.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub enum PropertyValue {
    /// String value.
    String(String),
    /// Integer value.
    Integer(i64),
    /// Float value.
    Float(f64),
    /// Boolean value.
    Boolean(bool),
    /// Date/time value (ISO 8601).
    DateTime(String),
    /// Binary blob (base64 encoded).
    Blob(String),
}
