//! Extract nodes from the IR by ID or type.

use anyhow::{bail, Result};
use docir_app::ParserConfig;
use docir_core::ir::IRNode;
use docir_core::types::NodeId;
use serde::Serialize;
use std::path::PathBuf;

use crate::commands::util::{parse_document, parse_node_type, write_json_output};

#[derive(Debug, Serialize)]
struct ExtractResult {
    nodes: Vec<IRNode>,
}

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

    let parsed = parse_document(&input, parser_config)?;

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

    let result = ExtractResult { nodes };
    write_json_output(&result, pretty, output)
}

fn parse_node_id(input: &str) -> Result<NodeId> {
    let trimmed = input.trim();
    let num = if let Some(hex) = trimmed.strip_prefix("node_") {
        u64::from_str_radix(hex, 16).map_err(|_| anyhow::anyhow!("Invalid node id: {input}"))?
    } else {
        trimmed
            .parse::<u64>()
            .map_err(|_| anyhow::anyhow!("Invalid node id: {input}"))?
    };
    Ok(NodeId::from_raw(num))
}
