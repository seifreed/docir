//! File-system adapter for reading document files.
//!
//! These functions are infrastructure adapters that bridge filesystem IO
//! to the application's byte-level functions. They belong at the edges
//! of the architecture, not in the domain or application core.

use crate::AppResult;
use docir_parser::ParseError as ParserParseError;
use std::fs::File;
use std::io::Read;
use std::path::Path;

pub(crate) fn read_bounded_file<P: AsRef<Path>>(
    path: P,
    max_input_size: u64,
) -> AppResult<Vec<u8>> {
    let file = File::open(path.as_ref()).map_err(ParserParseError::from)?;
    let read_limit = max_input_size.saturating_add(1);
    let initial_capacity = read_limit.min(64 * 1024).min(usize::MAX as u64) as usize;
    let mut data = Vec::with_capacity(initial_capacity);
    file.take(read_limit)
        .read_to_end(&mut data)
        .map_err(ParserParseError::from)?;
    if data.len() as u64 > max_input_size {
        return Err(ParserParseError::ResourceLimit(format!(
            "Input too large: {} bytes (max: {} bytes)",
            data.len(),
            max_input_size
        ))
        .into());
    }
    Ok(data)
}

/// Reads a file with size bounding, then delegates to a byte-level function.
/// Eliminates the repeated _path/_bytes boilerplate across inspection modules.
pub(crate) fn with_file_bytes<T>(
    path: impl AsRef<Path>,
    max_input_size: u64,
    f: impl FnOnce(&[u8]) -> AppResult<T>,
) -> AppResult<T> {
    let bytes = read_bounded_file(path, max_input_size)?;
    f(&bytes)
}

/// Reads a file with size bounding, then delegates to a byte-level function
/// that also needs the parser config.
pub(crate) fn with_file_bytes_and_config<T>(
    path: impl AsRef<Path>,
    config: &crate::config::ParserConfig,
    f: impl FnOnce(&[u8], &crate::config::ParserConfig) -> AppResult<T>,
) -> AppResult<T> {
    let bytes = read_bounded_file(path, config.max_input_size)?;
    f(&bytes, config)
}

#[cfg(test)]
mod tests {
    use super::read_bounded_file;
    use std::fs;

    #[test]
    fn read_bounded_file_rejects_content_above_limit() {
        let path = std::env::temp_dir().join(format!(
            "docir-app-bounded-read-{}-{}.bin",
            std::process::id(),
            std::thread::current().name().unwrap_or("test")
        ));
        fs::write(&path, b"0123456789").expect("write fixture");

        let result = read_bounded_file(&path, 4);

        fs::remove_file(&path).expect("remove fixture");
        let error = result.expect_err("oversized content must be rejected");
        assert!(error.to_string().contains("Input too large"));
    }
}
