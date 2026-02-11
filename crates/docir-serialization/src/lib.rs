//! # docir-serialization
//!
//! IR serialization for docir. Provides JSON and other output formats.

pub mod json;

pub use json::JsonSerializer;

use docir_core::ir::IRNode;
use docir_core::visitor::IrStore;
use docir_core::NodeId;
use thiserror::Error;

/// Serialization errors.
#[derive(Debug, Error)]
pub enum SerializationError {
    #[error("JSON error: {0}")]
    Json(#[from] serde_json::Error),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Node not found: {0}")]
    NodeNotFound(String),
}

/// Trait for IR serializers.
pub trait IrSerializer {
    /// Serializes a single node.
    fn serialize_node(&self, node: &IRNode) -> Result<Vec<u8>, SerializationError>;

    /// Serializes a complete IR tree.
    fn serialize_tree(
        &self,
        store: &IrStore,
        root_id: NodeId,
    ) -> Result<Vec<u8>, SerializationError>;

    /// Serializes to a string (for text formats).
    fn serialize_to_string(
        &self,
        store: &IrStore,
        root_id: NodeId,
    ) -> Result<String, SerializationError>;
}
