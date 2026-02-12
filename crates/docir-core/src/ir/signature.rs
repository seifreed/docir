//! Digital signature nodes.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Digital signature information.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct DigitalSignature {
    /// Unique identifier for this node.
    pub id: NodeId,
    /// Signature Id attribute, if present.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub signature_id: Option<String>,
    /// Signature method algorithm URI.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub signature_method: Option<String>,
    /// Digest method algorithm URIs.
    pub digest_methods: Vec<String>,
    /// Signer name or subject, if present.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub signer: Option<String>,
    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl DigitalSignature {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            signature_id: None,
            signature_method: None,
            digest_methods: Vec::new(),
            signer: None,
            span: None,
        }
    }
}
