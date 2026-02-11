//! Comments, footnotes, endnotes, headers, footers.

use crate::types::{NodeId, SourceSpan};
use serde::{Deserialize, Serialize};

/// Comment node.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Comment {
    pub id: NodeId,
    pub comment_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub initials: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub para_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub done: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    pub content: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Comment {
    pub fn new(comment_id: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            comment_id: comment_id.into(),
            author: None,
            initials: None,
            parent_id: None,
            para_id: None,
            done: None,
            date: None,
            content: Vec::new(),
            span: None,
        }
    }
}

/// Footnote node.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Footnote {
    pub id: NodeId,
    pub footnote_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub note_type: Option<String>,
    pub content: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Footnote {
    pub fn new(footnote_id: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            footnote_id: footnote_id.into(),
            note_type: None,
            content: Vec::new(),
            span: None,
        }
    }
}

/// Endnote node.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Endnote {
    pub id: NodeId,
    pub endnote_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub note_type: Option<String>,
    pub content: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Endnote {
    pub fn new(endnote_id: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            endnote_id: endnote_id.into(),
            note_type: None,
            content: Vec::new(),
            span: None,
        }
    }
}

/// Header node.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Header {
    pub id: NodeId,
    pub content: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Header {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            content: Vec::new(),
            span: None,
        }
    }
}

/// Footer node.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Footer {
    pub id: NodeId,
    pub content: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Footer {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            content: Vec::new(),
            span: None,
        }
    }
}

/// Comment range start marker.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentRangeStart {
    pub id: NodeId,
    pub comment_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl CommentRangeStart {
    pub fn new(comment_id: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            comment_id: comment_id.into(),
            span: None,
        }
    }
}

/// Comment range end marker.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentRangeEnd {
    pub id: NodeId,
    pub comment_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl CommentRangeEnd {
    pub fn new(comment_id: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            comment_id: comment_id.into(),
            span: None,
        }
    }
}

/// Comment reference marker (inline).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentReference {
    pub id: NodeId,
    pub comment_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl CommentReference {
    pub fn new(comment_id: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            comment_id: comment_id.into(),
            span: None,
        }
    }
}
