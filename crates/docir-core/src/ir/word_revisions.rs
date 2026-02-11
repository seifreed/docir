//! Word revisions (track changes) and comment metadata parts.

use crate::types::{NodeId, SourceSpan};
use serde::{Deserialize, Serialize};

/// A tracked change (insert/delete).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Revision {
    pub id: NodeId,
    pub change_type: RevisionType,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub revision_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    pub content: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Revision {
    pub fn new(change_type: RevisionType) -> Self {
        Self {
            id: NodeId::new(),
            change_type,
            revision_id: None,
            author: None,
            date: None,
            content: Vec::new(),
            span: None,
        }
    }
}

/// Revision type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum RevisionType {
    Insert,
    Delete,
    MoveFrom,
    MoveTo,
    FormatChange,
}

/// Comments extended metadata set (commentsExtended.xml).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentExtensionSet {
    pub id: NodeId,
    pub entries: Vec<CommentExtension>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl CommentExtensionSet {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            entries: Vec::new(),
            span: None,
        }
    }
}

/// A single comment extension entry.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentExtension {
    pub comment_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub para_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent_para_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub done: Option<bool>,
}

/// Comment ID mapping (commentsIds.xml).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentIdMap {
    pub id: NodeId,
    pub mappings: Vec<CommentIdMapEntry>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl CommentIdMap {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            mappings: Vec::new(),
            span: None,
        }
    }
}

/// Mapping entry between comment IDs.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentIdMapEntry {
    pub comment_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub para_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent_para_id: Option<String>,
}
