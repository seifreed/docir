//! Grep-like semantic search (text contains).

use anyhow::{bail, Result};
use docir_app::ParserConfig;
use std::path::PathBuf;

use crate::commands::query::{run_with_filters, QueryFilters};

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
    run_with_filters(
        input,
        QueryFilters {
            node_type,
            contains: Some(pattern),
            format,
            has_external_refs: None,
            has_macros: None,
        },
        pretty,
        output,
        parser_config,
    )
}
