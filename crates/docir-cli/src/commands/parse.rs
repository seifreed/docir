//! Parse command implementation.

use anyhow::{Context, Result};
use docir_app::ParserConfig;
use std::path::PathBuf;

use crate::commands::util::{build_app_and_parse, write_text_output};
use crate::OutputFormat;

pub fn run(
    input: PathBuf,
    format: OutputFormat,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let (app, parsed) = build_app_and_parse(&input, parser_config)?;

    // Serialize based on format
    let output_data = match format {
        OutputFormat::Json => app
            .serialize_json(&parsed, pretty)
            .context("Failed to serialize to JSON")?,
    };

    let output_path = output.clone();
    write_text_output(&output_data, output)?;
    if let Some(path) = output_path {
        eprintln!("Output written to {}", path.display());
    }
    Ok(())
}
