//! Spreadsheet slicers and timelines.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Slicer part (xl/slicers/slicer*.xml).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SlicerPart {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub caption: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cache_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub target_ref: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl SlicerPart {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            caption: None,
            cache_id: None,
            target_ref: None,
            span: None,
        }
    }
}

/// Timeline part (xl/timelines/timeline*.xml).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct TimelinePart {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cache_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl TimelinePart {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            cache_id: None,
            span: None,
        }
    }
}
