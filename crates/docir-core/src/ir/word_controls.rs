//! Word content controls, bookmarks, and fields.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Content control (SDT) node.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ContentControl {
    pub id: NodeId,
    pub content: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tag: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alias: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sdt_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub control_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_binding_xpath: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_binding_store_item_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_binding_prefix_mappings: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl ContentControl {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            content: Vec::new(),
            tag: None,
            alias: None,
            sdt_id: None,
            control_type: None,
            data_binding_xpath: None,
            data_binding_store_item_id: None,
            data_binding_prefix_mappings: None,
            span: None,
        }
    }
}

/// Bookmark start marker.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct BookmarkStart {
    pub id: NodeId,
    pub bookmark_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub col_first: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub col_last: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl BookmarkStart {
    pub fn new(bookmark_id: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            bookmark_id: bookmark_id.into(),
            name: None,
            col_first: None,
            col_last: None,
            span: None,
        }
    }
}

/// Bookmark end marker.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct BookmarkEnd {
    pub id: NodeId,
    pub bookmark_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl BookmarkEnd {
    pub fn new(bookmark_id: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            bookmark_id: bookmark_id.into(),
            span: None,
        }
    }
}

/// Field node (simple field or aggregated instructions).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Field {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub instruction: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub instruction_parsed: Option<FieldInstruction>,
    pub runs: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Field {
    pub fn new(instruction: Option<String>) -> Self {
        Self {
            id: NodeId::new(),
            instruction,
            instruction_parsed: None,
            runs: Vec::new(),
            span: None,
        }
    }
}

/// Parsed field instruction (subset).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct FieldInstruction {
    pub kind: FieldKind,
    #[serde(default)]
    pub args: Vec<String>,
    #[serde(default)]
    pub switches: Vec<String>,
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub enum FieldKind {
    Hyperlink,
    IncludeText,
    IncludePicture,
    MergeField,
    Date,
    Ref,
    PageRef,
    FootnoteRef,
    EndnoteRef,
    Dde,
    DdeAuto,
    AutoText,
    AutoCorrect,
    Unknown,
}
