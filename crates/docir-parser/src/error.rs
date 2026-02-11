//! Parser error types.

use thiserror::Error;

/// Errors that can occur during OOXML parsing.
#[derive(Debug, Error)]
pub enum ParseError {
    /// I/O error reading file.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    /// Invalid ZIP archive structure.
    #[error("Invalid ZIP structure: {0}")]
    InvalidZip(String),

    /// ZIP bomb or resource exhaustion detected.
    #[error("Resource limit exceeded: {0}")]
    ResourceLimit(String),

    /// XML parsing error.
    #[error("XML error in {file}: {message}")]
    Xml { file: String, message: String },

    /// Missing required OOXML part.
    #[error("Missing required part: {0}")]
    MissingPart(String),

    /// Invalid OOXML structure.
    #[error("Invalid OOXML structure: {0}")]
    InvalidStructure(String),

    /// Invalid format or unexpected layout.
    #[error("Invalid format: {0}")]
    InvalidFormat(String),

    /// Unsupported document format.
    #[error("Unsupported format: {0}")]
    UnsupportedFormat(String),

    /// Content type mismatch.
    #[error("Content type mismatch: expected {expected}, got {actual}")]
    ContentTypeMismatch { expected: String, actual: String },

    /// Relationship not found.
    #[error("Relationship not found: {0}")]
    RelationshipNotFound(String),

    /// Encoding error.
    #[error("Encoding error: {0}")]
    Encoding(String),

    /// Path traversal attempt detected.
    #[error("Path traversal attempt detected: {0}")]
    PathTraversal(String),
}

impl From<zip::result::ZipError> for ParseError {
    fn from(err: zip::result::ZipError) -> Self {
        ParseError::InvalidZip(err.to_string())
    }
}

impl From<quick_xml::Error> for ParseError {
    fn from(err: quick_xml::Error) -> Self {
        ParseError::Xml {
            file: String::new(),
            message: err.to_string(),
        }
    }
}
