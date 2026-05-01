//! DrawingML parts (word/drawings/*.xml).

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// DrawingML part containing shapes/images.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct DrawingPart {
    pub id: NodeId,
    pub path: String,
    pub shapes: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl DrawingPart {
    /// Public API entrypoint: new.
    pub fn new(path: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            path: path.into(),
            shapes: Vec::new(),
            span: None,
        }
    }
}
