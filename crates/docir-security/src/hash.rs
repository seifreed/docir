//! Hash helpers for security-relevant artifact fingerprints.

use sha2::{Digest, Sha256};

/// Computes a lowercase SHA-256 hex digest for a byte slice.
pub fn sha256_hex(data: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(data);
    format!("{:x}", hasher.finalize())
}
