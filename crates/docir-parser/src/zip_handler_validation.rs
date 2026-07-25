use crate::error::ParseError;

use super::ZipConfig;

pub(super) fn validate_archive_entry(
    name: &str,
    uncompressed_size: u64,
    compressed_size: u64,
    config: &ZipConfig,
) -> Result<(), ParseError> {
    if is_path_traversal(name) {
        return Err(ParseError::PathTraversal(name.to_string()));
    }

    let depth = name.matches('/').count();
    if depth > config.max_path_depth {
        return Err(ParseError::ResourceLimit(format!(
            "Path too deep: {} (max depth: {})",
            name, config.max_path_depth
        )));
    }

    if uncompressed_size > config.max_file_size {
        return Err(ParseError::ResourceLimit(format!(
            "File too large: {} ({} bytes, max: {} bytes)",
            name, uncompressed_size, config.max_file_size
        )));
    }

    if compressed_size > 0 {
        let ratio = uncompressed_size as f64 / compressed_size as f64;
        if ratio > config.max_compression_ratio {
            return Err(ParseError::ResourceLimit(format!(
                "Suspicious compression ratio for {}: {:.1}:1 (max: {:.1}:1)",
                name, ratio, config.max_compression_ratio
            )));
        }
    }

    Ok(())
}

pub(super) fn validate_total_size(
    total_uncompressed: u64,
    config: &ZipConfig,
) -> Result<(), ParseError> {
    if total_uncompressed > config.max_total_size {
        return Err(ParseError::ResourceLimit(format!(
            "Total uncompressed size too large: {} bytes (max: {} bytes)",
            total_uncompressed, config.max_total_size
        )));
    }
    Ok(())
}

pub(super) fn is_path_traversal(path: &str) -> bool {
    if path.contains("..") {
        return true;
    }

    let lower = path.to_ascii_lowercase();
    if lower.contains("%2e") || lower.contains("%2f") || lower.contains("%5c") {
        return true;
    }

    if path.contains('\0') {
        return true;
    }

    if path.starts_with('/') || path.starts_with('\\') {
        return true;
    }

    let bytes = path.as_bytes();
    if bytes.len() >= 2 {
        let first = bytes[0];
        let second = bytes[1];
        if first.is_ascii_alphabetic() && second == b':' {
            return true;
        }
    }

    path.contains('\\')
}
