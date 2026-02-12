//! Numbering IR nodes.

use crate::ir::{StyleParagraphProperties, StyleRunProperties, TextAlignment};
use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Numbering definitions.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct NumberingSet {
    pub id: NodeId,
    pub abstract_nums: Vec<AbstractNum>,
    pub nums: Vec<NumInstance>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl NumberingSet {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            abstract_nums: Vec::new(),
            nums: Vec::new(),
            span: None,
        }
    }
}

/// Abstract numbering definition.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct AbstractNum {
    pub abstract_id: u32,
    pub levels: Vec<NumberingLevel>,
}

/// Numbering instance mapping.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct NumInstance {
    pub num_id: u32,
    pub abstract_id: u32,
}

/// Numbering level.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct NumberingLevel {
    pub level: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<TextAlignment>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub suffix: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paragraph_props: Option<StyleParagraphProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub run_props: Option<StyleRunProperties>,
}
