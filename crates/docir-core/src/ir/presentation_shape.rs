use crate::ir::new_node_id;
use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// A shape on a slide.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Shape {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Shape name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,

    /// Alternative text (if provided).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alt_text: Option<String>,

    /// Shape type.
    pub shape_type: ShapeType,

    /// Position and size.
    pub transform: ShapeTransform,

    /// Text content (if this is a text shape).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<ShapeText>,

    /// Hyperlink (if this shape is clickable).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink: Option<String>,

    /// Relationship ID for linked media (if any).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,

    /// Media target path (if any).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub media_target: Option<String>,

    /// Linked media asset node (if resolved).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub media_asset: Option<NodeId>,

    /// Linked chart node (if resolved).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub chart_id: Option<NodeId>,

    /// Linked SmartArt part nodes (if resolved).
    #[serde(default)]
    pub smartart_parts: Vec<NodeId>,

    /// Related target paths (diagram parts, embeds).
    #[serde(default)]
    pub related_targets: Vec<String>,

    /// Linked OLE object node (if resolved).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ole_object: Option<NodeId>,

    /// Table content attached to this shape (if any).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub table: Option<NodeId>,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Shape {
    /// Creates a new Shape.
    pub fn new(shape_type: ShapeType) -> Self {
        Self {
            id: new_node_id(),
            name: None,
            alt_text: None,
            shape_type,
            transform: ShapeTransform::default(),
            text: None,
            hyperlink: None,
            relationship_id: None,
            media_target: None,
            media_asset: None,
            chart_id: None,
            smartart_parts: Vec::new(),
            related_targets: Vec::new(),
            ole_object: None,
            table: None,
            span: None,
        }
    }
}

/// Shape types.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeType {
    /// Rectangle shape.
    Rectangle,
    /// Rounded rectangle.
    RoundRect,
    /// Ellipse/circle.
    Ellipse,
    /// Triangle.
    Triangle,
    /// Line.
    Line,
    /// Arrow.
    Arrow,
    /// Text box.
    TextBox,
    /// Picture/image placeholder.
    Picture,
    /// Chart.
    Chart,
    /// Table.
    Table,
    /// Audio.
    Audio,
    /// Video.
    Video,
    /// OLE embedded object.
    OleObject,
    /// Group of shapes.
    Group,
    /// Custom/freeform shape.
    Custom,
    /// Unknown shape type.
    Unknown,
}

/// Shape position and size.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct ShapeTransform {
    /// X offset from slide origin in EMUs (914400 EMUs = 1 inch).
    pub x: i64,
    /// Y offset from slide origin in EMUs.
    pub y: i64,
    /// Width in EMUs.
    pub width: u64,
    /// Height in EMUs.
    pub height: u64,
    /// Rotation in 60000ths of a degree.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<i32>,
    /// Horizontal flip.
    #[serde(default)]
    pub flip_h: bool,
    /// Vertical flip.
    #[serde(default)]
    pub flip_v: bool,
}

/// Text content within a shape.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ShapeText {
    /// Paragraphs of text.
    pub paragraphs: Vec<ShapeTextParagraph>,
}

/// A paragraph within shape text.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ShapeTextParagraph {
    /// Runs of text.
    pub runs: Vec<ShapeTextRun>,
    /// Paragraph alignment.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<crate::ir::TextAlignment>,
}

/// A run of text within a shape paragraph.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ShapeTextRun {
    /// Text content.
    pub text: String,
    /// Bold.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    /// Italic.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    /// Font size in hundredths of a point.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<u32>,
    /// Font family.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
}
