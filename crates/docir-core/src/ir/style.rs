//! Styles IR nodes.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Collection of styles.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct StyleSet {
    /// Unique identifier for this node.
    pub id: NodeId,
    /// Styles list.
    pub styles: Vec<Style>,
    /// True if this set comes from stylesWithEffects.xml.
    #[serde(default)]
    pub with_effects: bool,
    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl StyleSet {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            styles: Vec::new(),
            with_effects: false,
            span: None,
        }
    }

    /// Public API entrypoint: with_effects.
    pub fn with_effects() -> Self {
        Self {
            with_effects: true,
            ..Self::new()
        }
    }
}

/// A style definition.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Style {
    /// Style id.
    pub style_id: String,
    /// Style name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Style type.
    pub style_type: StyleType,
    /// Based-on style id.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub based_on: Option<String>,
    /// Next style id.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub next: Option<String>,
    /// Is default.
    #[serde(default)]
    pub is_default: bool,
    /// Run properties (subset).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub run_props: Option<StyleRunProperties>,
    /// Paragraph properties (subset).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paragraph_props: Option<StyleParagraphProperties>,
    /// Table properties (subset).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub table_props: Option<crate::ir::TableProperties>,
}

/// Style type.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum StyleType {
    Paragraph,
    Character,
    Table,
    Numbering,
    #[default]
    Other,
}

/// Run properties for styles.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct StyleRunProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<crate::ir::UnderlineStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub strike: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub highlight: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<crate::ir::VerticalTextAlignment>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub all_caps: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub small_caps: Option<bool>,
}

/// Paragraph properties for styles.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct StyleParagraphProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<crate::ir::TextAlignment>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indentation: Option<crate::ir::Indentation>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spacing: Option<crate::ir::Spacing>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub numbering: Option<crate::ir::NumberingInfo>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<crate::ir::ParagraphBorders>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub keep_next: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub keep_lines: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_break_before: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub widow_control: Option<bool>,
}

#[cfg(test)]
mod tests {
    use super::{StyleSet, StyleType};

    #[test]
    fn style_set_new_starts_without_effects() {
        let set = StyleSet::new();
        assert!(set.styles.is_empty());
        assert!(!set.with_effects);
        assert!(set.span.is_none());
    }

    #[test]
    fn style_set_with_effects_sets_flag() {
        let set = StyleSet::with_effects();
        assert!(set.with_effects);
        assert!(set.styles.is_empty());
    }

    #[test]
    fn style_type_default_is_other() {
        assert_eq!(StyleType::default(), StyleType::Other);
    }
}
