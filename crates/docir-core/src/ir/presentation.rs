//! Presentation (PPTX) IR nodes.

use crate::ir::new_node_id;
use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

#[path = "presentation_shape.rs"]
mod presentation_shape;
pub use presentation_shape::*;

fn new_presentation_node_id() -> NodeId {
    new_node_id()
}

fn collect_children_nodes(
    primary: &[NodeId],
    secondary: &[NodeId],
    trailing: Option<NodeId>,
) -> Vec<NodeId> {
    let mut out = primary.to_vec();
    out.extend(secondary.iter().copied());
    if let Some(id) = trailing {
        out.push(id);
    }
    out
}

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
            id: new_presentation_node_id(),
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
        collect_children_nodes(&self.shapes, &self.comments, self.notes_slide)
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
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: new_presentation_node_id(),
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
    #[serde(skip_serializing_if = "Option::is_none")]
    pub auto_compress_pictures: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub compat_mode: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtl: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
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
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: new_presentation_node_id(),
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
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: new_presentation_node_id(),
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
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: new_presentation_node_id(),
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
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: new_presentation_node_id(),
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
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: new_presentation_node_id(),
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
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: new_presentation_node_id(),
            name: None,
            preserve: None,
            show_master_sp: None,
            show_master_ph_anim: None,
            shapes: Vec::new(),
            layouts: Vec::new(),
            span: None,
        }
    }

    /// Public API entrypoint: children.
    pub fn children(&self) -> Vec<NodeId> {
        collect_children_nodes(&self.shapes, &self.layouts, None)
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
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: new_presentation_node_id(),
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

    /// Public API entrypoint: children.
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
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: new_presentation_node_id(),
            name: None,
            shapes: Vec::new(),
            span: None,
        }
    }

    /// Public API entrypoint: children.
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
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: new_presentation_node_id(),
            name: None,
            shapes: Vec::new(),
            span: None,
        }
    }

    /// Public API entrypoint: children.
    pub fn children(&self) -> Vec<NodeId> {
        self.shapes.clone()
    }
}
