//! Parse command implementation.

use anyhow::{Context, Result};
use docir_parser::ParserConfig;
use docir_serialization::json::to_json;
use std::fs;
use std::io::Write;
use std::path::PathBuf;

use crate::commands::util::build_parser;
use crate::OutputFormat;

pub fn run(
    input: PathBuf,
    format: OutputFormat,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    // Parse the document
    let parser = build_parser(parser_config);
    let parsed = parser
        .parse_file(&input)
        .with_context(|| format!("Failed to parse {}", input.display()))?;

    // Serialize based on format
    let output_data = match format {
        OutputFormat::Json => {
            to_json(&parsed.store, parsed.root_id, pretty).context("Failed to serialize to JSON")?
        }
    };

    // Write output
    if let Some(output_path) = output {
        fs::write(&output_path, &output_data)
            .with_context(|| format!("Failed to write to {}", output_path.display()))?;
        eprintln!("Output written to {}", output_path.display());
    } else {
        let stdout = std::io::stdout();
        let mut handle = stdout.lock();
        handle.write_all(output_data.as_bytes())?;
        handle.write_all(b"\n")?;
    }

    Ok(())
}
