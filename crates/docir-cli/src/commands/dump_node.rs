//! Dump node command implementation.

use anyhow::{Context, Result};
use docir_app::ParserConfig;
use std::path::PathBuf;

use crate::commands::util::{parse_document, parse_node_id};
use crate::OutputFormat;

pub fn run(
    input: PathBuf,
    node_id_str: &str,
    format: OutputFormat,
    parser_config: &ParserConfig,
) -> Result<()> {
    let parsed = parse_document(&input, parser_config)?;

    // Parse the node ID ("node_XXXXXXXX" hex or raw number)
    let node_id = parse_node_id(node_id_str)
        .with_context(|| format!("Invalid node ID format: {}", node_id_str))?;

    // Find the node
    let node = parsed
        .store()
        .get(node_id)
        .ok_or_else(|| anyhow::anyhow!("Node not found: {}", node_id))?;

    // Serialize based on format
    match format {
        OutputFormat::Json => {
            let json = serde_json::to_string_pretty(node)?;
            println!("{}", json);
        }
    }

    Ok(())
}
