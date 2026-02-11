//! VML legacy drawings.

use crate::types::{NodeId, SourceSpan};
use serde::{Deserialize, Serialize};

/// VML drawing part (word/vmlDrawing*.vml).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct VmlDrawing {
    pub id: NodeId,
    pub path: String,
    pub shapes: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl VmlDrawing {
    pub fn new(path: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            path: path.into(),
            shapes: Vec::new(),
            span: None,
        }
    }
}

/// VML shape (legacy).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct VmlShape {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shape_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub filled: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroked: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rel_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub image_target: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl VmlShape {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            shape_type: None,
            style: None,
            filled: None,
            stroked: None,
            rel_id: None,
            image_target: None,
            text: None,
            span: None,
        }
    }
}
