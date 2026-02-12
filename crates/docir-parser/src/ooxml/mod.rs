//! OOXML-specific parsing modules.

pub mod content_types;
pub mod docx;
pub mod part_registry;
pub mod part_utils;
pub mod pptx;
pub mod relationships;
pub mod shared;
pub mod xlsx;
mod xml_utils;

pub use content_types::ContentTypes;
pub use relationships::{Relationship, Relationships};
