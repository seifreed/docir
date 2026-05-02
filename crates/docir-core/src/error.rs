//! Core error types for docir.

use thiserror::Error;

/// Core errors that can occur in the IR layer.
#[derive(Debug, Error)]
pub enum CoreError {
    /// Node not found in the IR tree.
    #[error("Node not found: {0}")]
    NodeNotFound(String),

    /// Invalid node type for the requested operation.
    #[error("Invalid node type: expected {expected}, got {actual}")]
    InvalidNodeType { expected: String, actual: String },

    /// Invalid node reference.
    #[error("Invalid node reference: {0}")]
    InvalidReference(String),

    /// Visitor error during traversal.
    #[error("Visitor error: {0}")]
    VisitorError(#[source] Box<dyn std::error::Error + Send + Sync>),
}
