//! Presentation (PPTX) IR nodes.

use crate::ir::new_node_id;
use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

#[path = "presentation_shape.rs"]
mod presentation_shape;
pub use presentation_shape::*;
#[path = "presentation_parts.rs"]
mod presentation_parts;
pub use presentation_parts::*;

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
