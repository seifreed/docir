//! Summary command implementation.

use anyhow::{Context, Result};
use docir_app::ParserConfig;
use std::path::PathBuf;

use crate::commands::util::{build_app_and_parse, write_text_output};

/// Public API entrypoint: run.
pub fn run(input: PathBuf, parser_config: &ParserConfig) -> Result<()> {
    let (app, parsed) = build_app_and_parse(&input, parser_config)?;
    let source = input.to_string_lossy();
    let summary_text = app
        .format_summary(&parsed, Some(&source))
        .context("Failed to build document summary")?;

    write_text_output(&summary_text, None)
}
