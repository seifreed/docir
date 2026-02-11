//! DrawingML parts (word/drawings/*.xml).

use crate::types::{NodeId, SourceSpan};
use serde::{Deserialize, Serialize};

/// DrawingML part containing shapes/images.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DrawingPart {
    pub id: NodeId,
    pub path: String,
    pub shapes: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl DrawingPart {
    pub fn new(path: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            path: path.into(),
            shapes: Vec::new(),
            span: None,
        }
    }
}
