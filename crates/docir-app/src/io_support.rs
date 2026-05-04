//! File-system adapter for reading document files.
//!
//! These functions are infrastructure adapters that bridge filesystem IO
//! to the application's byte-level functions. They belong at the edges
//! of the architecture, not in the domain or application core.

use crate::AppResult;
use docir_parser::ParseError as ParserParseError;
use std::fs;
use std::path::Path;

pub(crate) fn read_bounded_file<P: AsRef<Path>>(
    path: P,
    max_input_size: u64,
) -> AppResult<Vec<u8>> {
    let path = path.as_ref();
    let metadata = fs::metadata(path).map_err(ParserParseError::from)?;
    if metadata.len() > max_input_size {
        return Err(ParserParseError::ResourceLimit(format!(
            "Input exceeds max_input_size ({} > {})",
            metadata.len(),
            max_input_size
        ))
        .into());
    }
    fs::read(path)
        .map_err(ParserParseError::from)
        .map_err(Into::into)
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
