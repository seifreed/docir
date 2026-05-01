//! Theme IR nodes.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Document theme.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Theme {
    /// Unique identifier for this node.
    pub id: NodeId,
    /// Theme name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Color scheme.
    pub colors: Vec<ThemeColor>,
    /// Font scheme.
    pub fonts: ThemeFontScheme,
    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Theme {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            colors: Vec::new(),
            fonts: ThemeFontScheme::default(),
            span: None,
        }
    }
}

/// Theme color entry.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ThemeColor {
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
}

/// Theme font scheme.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct ThemeFontScheme {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub major: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub minor: Option<String>,
}
