//! Dump node command implementation.

use anyhow::{Context, Result};
use docir_app::ParserConfig;
use std::path::PathBuf;

use crate::commands::util::{parse_document, parse_node_id, write_json_output};
use crate::OutputFormat;

/// Public API entrypoint: run.
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
            write_json_output(node, true, None)?;
        }
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::run;
    use crate::commands::util::parse_document;
    use crate::test_support;
    use crate::OutputFormat;
    use docir_app::ParserConfig;

    #[test]
    fn dump_node_run_reports_missing_node_for_valid_but_stale_id() {
        let input = test_support::fixture("minimal.docx");
        let parsed = parse_document(&input, &ParserConfig::default()).expect("parse document");
        let node_id = parsed.root_id().to_string();
        let err = run(
            input,
            &node_id,
            OutputFormat::Json,
            &ParserConfig::default(),
        )
        .expect_err("stale node id should not exist in a new parse");
        assert!(err.to_string().contains("Node not found"));
    }

    #[test]
    fn dump_node_run_rejects_invalid_id() {
        let err = run(
            test_support::fixture("minimal.docx"),
            "not-a-node-id",
            OutputFormat::Json,
            &ParserConfig::default(),
        )
        .expect_err("invalid id should fail");
        assert!(err.to_string().contains("Invalid node ID format"));
    }
}
