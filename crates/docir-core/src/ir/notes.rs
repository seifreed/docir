//! Comments, footnotes, endnotes, headers, footers.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Comment node.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
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
    /// Public API entrypoint: new.
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
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
    /// Public API entrypoint: new.
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
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
    /// Public API entrypoint: new.
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Header {
    pub id: NodeId,
    pub content: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Header {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            content: Vec::new(),
            span: None,
        }
    }
}

/// Footer node.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Footer {
    pub id: NodeId,
    pub content: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Footer {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            content: Vec::new(),
            span: None,
        }
    }
}

/// Comment range start marker.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct CommentRangeStart {
    pub id: NodeId,
    pub comment_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl CommentRangeStart {
    /// Public API entrypoint: new.
    pub fn new(comment_id: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            comment_id: comment_id.into(),
            span: None,
        }
    }
}

/// Comment range end marker.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct CommentRangeEnd {
    pub id: NodeId,
    pub comment_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl CommentRangeEnd {
    /// Public API entrypoint: new.
    pub fn new(comment_id: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            comment_id: comment_id.into(),
            span: None,
        }
    }
}

/// Comment reference marker (inline).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct CommentReference {
    pub id: NodeId,
    pub comment_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl CommentReference {
    /// Public API entrypoint: new.
    pub fn new(comment_id: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            comment_id: comment_id.into(),
            span: None,
        }
    }
}
