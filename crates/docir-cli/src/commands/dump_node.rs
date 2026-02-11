//! Dump node command implementation.

use anyhow::{bail, Context, Result};
use docir_core::NodeId;
use docir_parser::ParserConfig;
use std::path::PathBuf;

use crate::commands::util::build_app;
use crate::OutputFormat;

pub fn run(
    input: PathBuf,
    node_id_str: &str,
    format: OutputFormat,
    parser_config: &ParserConfig,
) -> Result<()> {
    // Parse the document
    let app = build_app(parser_config);
    let parsed = app
        .parse_file(&input)
        .with_context(|| format!("Failed to parse {}", input.display()))?;

    // Parse the node ID
    // Expected format: "node_XXXXXXXX" where X is hex
    let node_id = if node_id_str.starts_with("node_") {
        let hex_part = node_id_str.strip_prefix("node_").unwrap();
        let id_value = u64::from_str_radix(hex_part, 16)
            .with_context(|| format!("Invalid node ID format: {}", node_id_str))?;
        NodeId::from_raw(id_value)
    } else {
        // Try parsing as raw number
        let id_value: u64 = node_id_str
            .parse()
            .with_context(|| format!("Invalid node ID format: {}", node_id_str))?;
        NodeId::from_raw(id_value)
    };

    // Find the node
    let node = parsed
        .store
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
