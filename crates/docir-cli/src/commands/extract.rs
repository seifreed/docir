//! Extract nodes from the IR by ID or type.

use anyhow::{bail, Result};
use docir_app::ParserConfig;
use docir_core::ir::IRNode;
use serde::Serialize;
use std::path::PathBuf;

use crate::commands::util::{parse_node_id, parse_node_type, run_json_document_command};

#[derive(Debug, Serialize)]
struct ExtractResult {
    nodes: Vec<IRNode>,
}

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    node_ids: Vec<String>,
    node_type: Option<String>,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    if node_ids.is_empty() && node_type.is_none() {
        bail!("Provide --node-id or --node-type");
    }

    run_json_document_command(input, parser_config, pretty, output, move |parsed| {
        let mut nodes = Vec::new();

        for id in node_ids {
            let node_id = parse_node_id(&id)?;
            if let Some(node) = parsed.store().get(node_id) {
                nodes.push(node.clone());
            }
        }

        if let Some(t) = node_type {
            let node_type = parse_node_type(&t)?;
            for id in parsed.store().iter_ids_by_type(node_type) {
                if let Some(node) = parsed.store().get(id) {
                    nodes.push(node.clone());
                }
            }
        }

        Ok(ExtractResult { nodes })
    })
}

#[cfg(test)]
mod tests {
    use super::run;
    use docir_app::ParserConfig;
    use std::fs;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn fixture(name: &str) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../fixtures/ooxml")
            .join(name)
    }

    fn temp_file(name: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_extract_{name}_{nanos}.json"))
    }

    #[test]
    fn extract_run_requires_selector() {
        let err = run(
            fixture("minimal.docx"),
            Vec::new(),
            None,
            false,
            None,
            &ParserConfig::default(),
        )
        .expect_err("selectorless extract should fail");
        assert!(err.to_string().contains("Provide --node-id or --node-type"));
    }

    #[test]
    fn extract_run_with_node_type_writes_json() {
        let output = temp_file("by_type");
        run(
            fixture("minimal.docx"),
            Vec::new(),
            Some("Paragraph".to_string()),
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("extract by type");
        let text = fs::read_to_string(&output).expect("extract output");
        assert!(text.contains("nodes"));
        let _ = fs::remove_file(output);
    }
}
