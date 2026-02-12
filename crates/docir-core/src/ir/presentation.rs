//! Presentation (PPTX) IR nodes.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// A presentation slide.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Slide {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Slide number (1-based).
    pub number: u32,

    /// Slide name/title.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,

    /// Shapes on this slide.
    pub shapes: Vec<NodeId>,

    /// Comments on this slide.
    #[serde(default)]
    pub comments: Vec<NodeId>,

    /// Notes for this slide.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub notes: Option<String>,

    /// Notes slide node (if present).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub notes_slide: Option<NodeId>,

    /// Slide layout reference.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub layout_id: Option<String>,

    /// Slide master reference.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub master_id: Option<String>,

    /// Is this slide hidden?
    #[serde(default)]
    pub hidden: bool,

    /// Slide transition info.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub transition: Option<SlideTransition>,

    /// Animation entries (timing) for the slide.
    #[serde(default)]
    pub animations: Vec<SlideAnimation>,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Slide {
    /// Creates a new Slide with the given number.
    pub fn new(number: u32) -> Self {
        Self {
            id: NodeId::new(),
            number,
            name: None,
            shapes: Vec::new(),
            comments: Vec::new(),
            notes: None,
            notes_slide: None,
            layout_id: None,
            master_id: None,
            hidden: false,
            transition: None,
            animations: Vec::new(),
            span: None,
        }
    }

    /// Returns all child node IDs.
    pub fn children(&self) -> Vec<NodeId> {
        let mut out = self.shapes.clone();
        out.extend(self.comments.iter().copied());
        if let Some(notes_id) = self.notes_slide {
            out.push(notes_id);
        }
        out
    }
}

/// Slide transition settings.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SlideTransition {
    pub transition_type: Option<String>,
    pub speed: Option<String>,
    pub advance_on_click: Option<bool>,
    pub advance_after_ms: Option<u32>,
    pub duration_ms: Option<u32>,
}

/// Notes slide (ppt/notesSlides/notesSlide*.xml).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct NotesSlide {
    pub id: NodeId,
    pub shapes: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl NotesSlide {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            shapes: Vec::new(),
            text: None,
            span: None,
        }
    }
}

/// Slide animation entry.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SlideAnimation {
    pub animation_type: String,
    pub target: Option<String>,
    pub duration_ms: Option<u32>,
    pub preset_id: Option<String>,
    pub preset_class: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub media_asset: Option<NodeId>,
}

/// Presentation properties (presProps.xml).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct PresentationProperties {
    pub id: NodeId,
    pub auto_compress_pictures: Option<bool>,
    pub compat_mode: Option<String>,
    pub rtl: Option<bool>,
    pub show_special_placeholders: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub remove_personal_info_on_save: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_ink_annotation: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

/// Presentation info derived from presentation.xml.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct PresentationInfo {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub slide_size: Option<SlideSize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub notes_size: Option<SlideSize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_slide_num: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_loop: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_narration: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_animation: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub use_timings: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl PresentationInfo {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            slide_size: None,
            notes_size: None,
            first_slide_num: None,
            show_type: None,
            show_loop: None,
            show_narration: None,
            show_animation: None,
            use_timings: None,
            span: None,
        }
    }
}

/// Slide size definition.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SlideSize {
    pub cx: u64,
    pub cy: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size_type: Option<String>,
}

impl PresentationProperties {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            auto_compress_pictures: None,
            compat_mode: None,
            rtl: None,
            show_special_placeholders: None,
            remove_personal_info_on_save: None,
            show_ink_annotation: None,
            span: None,
        }
    }
}

/// View properties (viewProps.xml).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ViewProperties {
    pub id: NodeId,
    pub last_view: Option<String>,
    pub zoom: Option<u32>,
    pub show_comments: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_hidden_slides: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_guides: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_grid: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_outline_icons: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl ViewProperties {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            last_view: None,
            zoom: None,
            show_comments: None,
            show_hidden_slides: None,
            show_guides: None,
            show_grid: None,
            show_outline_icons: None,
            span: None,
        }
    }
}

/// Table styles list (tableStyles.xml).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct TableStyleSet {
    pub id: NodeId,
    pub default_style_id: Option<String>,
    pub styles: Vec<TableStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl TableStyleSet {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            default_style_id: None,
            styles: Vec::new(),
            span: None,
        }
    }
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct TableStyle {
    pub style_id: String,
    pub name: Option<String>,
}

/// PPTX comment author.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct PptxCommentAuthor {
    pub id: NodeId,
    pub author_id: u32,
    pub name: Option<String>,
    pub initials: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

/// PPTX comment.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct PptxComment {
    pub id: NodeId,
    pub author_id: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author_initials: Option<String>,
    pub datetime: Option<String>,
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

/// Presentation tag entry.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct PresentationTag {
    pub id: NodeId,
    pub name: String,
    pub value: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

/// People part (coauthoring).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct PeoplePart {
    pub id: NodeId,
    pub people: Vec<PersonEntry>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl PeoplePart {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            people: Vec::new(),
            span: None,
        }
    }
}

/// Person entry in people.xml.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct PersonEntry {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub person_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub user_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub display_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub initials: Option<String>,
}

/// SmartArt part.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SmartArtPart {
    pub id: NodeId,
    pub kind: String,
    pub path: String,
    pub root_element: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub point_count: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub connection_count: Option<u32>,
    #[serde(default)]
    pub rel_ids: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

/// Slide master part.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SlideMaster {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub preserve: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_master_sp: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_master_ph_anim: Option<bool>,
    pub shapes: Vec<NodeId>,
    pub layouts: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl SlideMaster {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            preserve: None,
            show_master_sp: None,
            show_master_ph_anim: None,
            shapes: Vec::new(),
            layouts: Vec::new(),
            span: None,
        }
    }

    pub fn children(&self) -> Vec<NodeId> {
        let mut out = self.shapes.clone();
        out.extend(self.layouts.iter().copied());
        out
    }
}

/// Slide layout part.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SlideLayout {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub layout_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub matching_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub preserve: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_master_sp: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_master_ph_anim: Option<bool>,
    pub shapes: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl SlideLayout {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            layout_type: None,
            matching_name: None,
            preserve: None,
            show_master_sp: None,
            show_master_ph_anim: None,
            shapes: Vec::new(),
            span: None,
        }
    }

    pub fn children(&self) -> Vec<NodeId> {
        self.shapes.clone()
    }
}

/// Notes master part.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct NotesMaster {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    pub shapes: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl NotesMaster {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            shapes: Vec::new(),
            span: None,
        }
    }

    pub fn children(&self) -> Vec<NodeId> {
        self.shapes.clone()
    }
}

/// Handout master part.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct HandoutMaster {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    pub shapes: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl HandoutMaster {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            shapes: Vec::new(),
            span: None,
        }
    }

    pub fn children(&self) -> Vec<NodeId> {
        self.shapes.clone()
    }
}

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
            id: NodeId::new(),
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
