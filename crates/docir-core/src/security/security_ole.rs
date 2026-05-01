use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// ActiveX control.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ActiveXControl {
    /// Unique identifier for this node.
    pub id: NodeId,
    /// Control name (if provided).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// CLSID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub clsid: Option<String>,
    /// ProgID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prog_id: Option<String>,
    /// Control properties (raw key/value).
    pub properties: Vec<(String, String)>,
    /// Source span.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl ActiveXControl {
    /// Creates an empty ActiveX control container with a fresh identifier.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            clsid: None,
            prog_id: None,
            properties: Vec::new(),
            span: None,
        }
    }
}

/// Embedded OLE object.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct OleObject {
    /// Unique identifier for this node.
    pub id: NodeId,
    /// ProgID (e.g., "Excel.Sheet.12").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prog_id: Option<String>,
    /// Class ID (GUID).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub class_id: Option<String>,
    /// Object name/label.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Class name exposed by the object if available.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub class_name: Option<String>,
    /// Is this a linked object (vs embedded)?
    pub is_linked: bool,
    /// Link target (if linked).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub link_target: Option<String>,
    /// Size in bytes.
    pub size_bytes: u64,
    /// SHA-256 hash of object data (hex).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_hash: Option<String>,
    /// Relationship ID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,
    /// Embedded payload kind when the OLE object wraps another file.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub embedded_payload_kind: Option<String>,
    /// Embedded file name recovered from Ole10Native / Package metadata.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub embedded_file_name: Option<String>,
    /// Original source path stored by the object.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_path: Option<String>,
    /// Temporary path stored by the object.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub temp_path: Option<String>,
    /// Estimated MIME type of the embedded payload.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mime_type: Option<String>,
    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl OleObject {
    /// Creates a new OleObject.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            prog_id: None,
            class_id: None,
            name: None,
            class_name: None,
            is_linked: false,
            link_target: None,
            size_bytes: 0,
            data_hash: None,
            relationship_id: None,
            embedded_payload_kind: None,
            embedded_file_name: None,
            source_path: None,
            temp_path: None,
            mime_type: None,
            span: None,
        }
    }
}

impl Default for OleObject {
    fn default() -> Self {
        Self::new()
    }
}
