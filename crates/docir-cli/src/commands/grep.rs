//! Grep-like semantic search (text contains).

use anyhow::{bail, Result};
use docir_parser::ParserConfig;
use std::path::PathBuf;

pub fn run(
    input: PathBuf,
    pattern: String,
    node_type: Option<String>,
    format: Option<String>,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    if pattern.trim().is_empty() {
        bail!("Pattern must not be empty");
    }
    crate::commands::query::run(
        input,
        node_type,
        Some(pattern),
        format,
        None,
        None,
        pretty,
        output,
        parser_config,
    )
}
