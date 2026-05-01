//! Paragraph and Run IR nodes.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// A paragraph containing runs of text.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Paragraph {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Runs of text within this paragraph.
    pub runs: Vec<NodeId>,

    /// Style identifier (references a style definition).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,

    /// Paragraph properties.
    pub properties: ParagraphProperties,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Paragraph {
    /// Creates a new empty Paragraph.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            runs: Vec::new(),
            style_id: None,
            properties: ParagraphProperties::default(),
            span: None,
        }
    }

    /// Returns all child node IDs.
    pub fn children(&self) -> Vec<NodeId> {
        self.runs.clone()
    }

    /// Extracts all plain text from this paragraph's runs.
    /// Requires the IR store to resolve run node IDs into text.
    pub fn text_content(&self, store: &crate::visitor::IrStore) -> String {
        self.runs
            .iter()
            .filter_map(|id| {
                store.get(*id).and_then(|node| {
                    if let crate::ir::IRNode::Run(run) = node {
                        Some(run.text.as_str())
                    } else {
                        None
                    }
                })
            })
            .collect()
    }
}

impl Default for Paragraph {
    fn default() -> Self {
        Self::new()
    }
}

/// Paragraph formatting properties.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct ParagraphProperties {
    /// Horizontal alignment.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<TextAlignment>,

    /// Indentation in twips.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indentation: Option<Indentation>,

    /// Spacing before/after in twips.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spacing: Option<Spacing>,

    /// Outline level (0-9 for headings).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,

    /// Is this a list item?
    #[serde(skip_serializing_if = "Option::is_none")]
    pub numbering: Option<NumberingInfo>,

    /// Paragraph borders.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<ParagraphBorders>,

    /// Keep with next paragraph.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub keep_next: Option<bool>,

    /// Keep lines together.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub keep_lines: Option<bool>,

    /// Page break before this paragraph.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_break_before: Option<bool>,

    /// Widow control enabled.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub widow_control: Option<bool>,
}

/// Text alignment options.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAlignment {
    Left,
    Center,
    Right,
    Justify,
    Distribute,
}

/// Paragraph indentation.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct Indentation {
    /// Left indent in twips.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<i32>,
    /// Right indent in twips.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<i32>,
    /// First line indent in twips.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_line: Option<i32>,
    /// Hanging indent in twips.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hanging: Option<i32>,
}

/// Paragraph spacing.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct Spacing {
    /// Space before paragraph in twips.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub before: Option<u32>,
    /// Space after paragraph in twips.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub after: Option<u32>,
    /// Line spacing value.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line: Option<u32>,
    /// Line spacing rule.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_rule: Option<LineSpacingRule>,
}

/// Line spacing rules.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineSpacingRule {
    Auto,
    Exact,
    AtLeast,
}

/// Numbering/list information.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct NumberingInfo {
    /// Numbering definition ID.
    pub num_id: u32,
    /// Indentation level.
    pub level: u32,
    /// Optional numbering format (RTF/ODF/OOXML mapping).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub format: Option<String>,
}

/// Paragraph border definition.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct ParagraphBorders {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<crate::ir::Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<crate::ir::Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<crate::ir::Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<crate::ir::Border>,
}

/// A run of text with uniform formatting.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Run {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// The text content of this run.
    pub text: String,

    /// Run formatting properties.
    pub properties: RunProperties,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Run {
    /// Creates a new Run with the given text.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            text: text.into(),
            properties: RunProperties::default(),
            span: None,
        }
    }

    /// Creates a new Run with text and properties.
    pub fn with_properties(text: impl Into<String>, properties: RunProperties) -> Self {
        Self {
            id: NodeId::new(),
            text: text.into(),
            properties,
            span: None,
        }
    }
}

/// Run formatting properties.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct RunProperties {
    /// Character style ID (if any).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
    /// Font family name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,

    /// Font size in half-points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<u32>,

    /// Bold formatting.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,

    /// Italic formatting.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,

    /// Underline style.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<UnderlineStyle>,

    /// Strike-through.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub strike: Option<bool>,

    /// Text color (hex without #).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,

    /// Highlight/background color.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub highlight: Option<String>,

    /// Vertical alignment (superscript/subscript).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<VerticalTextAlignment>,

    /// All caps formatting.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub all_caps: Option<bool>,

    /// Small caps formatting.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub small_caps: Option<bool>,
}

/// Underline styles.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnderlineStyle {
    Single,
    Double,
    Thick,
    Dotted,
    Dashed,
    Wave,
    None,
}

/// Vertical text alignment.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalTextAlignment {
    Baseline,
    Superscript,
    Subscript,
}

/// A hyperlink wrapping one or more runs.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Target URL or anchor.
    pub target: String,

    /// Whether this is an external link.
    pub is_external: bool,

    /// Relationship ID in the OOXML package.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,

    /// Tooltip text.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,

    /// Child runs.
    pub runs: Vec<NodeId>,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Hyperlink {
    /// Creates a new Hyperlink.
    pub fn new(target: impl Into<String>, is_external: bool) -> Self {
        Self {
            id: NodeId::new(),
            target: target.into(),
            is_external,
            relationship_id: None,
            tooltip: None,
            runs: Vec::new(),
            span: None,
        }
    }

    /// Returns all child node IDs.
    pub fn children(&self) -> Vec<NodeId> {
        self.runs.clone()
    }
}
