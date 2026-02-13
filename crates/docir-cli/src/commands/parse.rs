//! Parse command implementation.

use anyhow::{Context, Result};
use docir_app::ParserConfig;
use std::path::PathBuf;

use crate::commands::util::{build_app, write_text_output};
use crate::OutputFormat;

pub fn run(
    input: PathBuf,
    format: OutputFormat,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    // Parse the document
    let app = build_app(parser_config);
    let parsed = app
        .parse_file(&input)
        .with_context(|| format!("Failed to parse {}", input.display()))?;

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
