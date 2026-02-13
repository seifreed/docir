//! Shared CLI helpers.

use anyhow::{anyhow, Context, Result};
use docir_app::ParserConfig;
use docir_app::{DocirApp, ParsedDocument};
use docir_core::types::{
    parse_document_format as parse_core_document_format, parse_node_type as parse_core_node_type,
    DocumentFormat, NodeId, NodeType,
};
use serde::Serialize;
use std::fs;
use std::fs::File;
use std::io::{self, Write};
use std::path::PathBuf;

pub fn parse_node_type(input: &str) -> Result<NodeType> {
    parse_core_node_type(input).map_err(|e| anyhow!(e))
}

pub fn parse_doc_format(input: &str) -> Result<DocumentFormat> {
    parse_core_document_format(input).map_err(|e| anyhow!(e))
}

pub fn parse_node_id(input: &str) -> Result<NodeId> {
    let trimmed = input.trim();
    let value = if let Some(hex) = trimmed.strip_prefix("node_") {
        u64::from_str_radix(hex, 16).map_err(|_| anyhow!("Invalid node id: {input}"))?
    } else {
        trimmed
            .parse::<u64>()
            .map_err(|_| anyhow!("Invalid node id: {input}"))?
    };
    Ok(NodeId::from_raw(value))
}

pub fn build_app(config: &ParserConfig) -> DocirApp {
    DocirApp::new(config.clone())
}

pub fn parse_document(input: &PathBuf, parser_config: &ParserConfig) -> Result<ParsedDocument> {
    let (_, parsed) = build_app_and_parse(input, parser_config)?;
    Ok(parsed)
}

pub fn build_app_and_parse(
    input: &PathBuf,
    parser_config: &ParserConfig,
) -> Result<(DocirApp, ParsedDocument)> {
    let app = build_app(parser_config);
    let parsed = parse_with_context(&app, input)?;
    Ok((app, parsed))
}

fn parse_with_context(app: &DocirApp, input: &PathBuf) -> Result<ParsedDocument> {
    app.parse_file(input)
        .with_context(|| format!("Failed to parse {}", input.display()))
}

pub fn write_json_output<T: Serialize>(
    value: &T,
    pretty: bool,
    output: Option<PathBuf>,
) -> Result<()> {
    let mut writer: Box<dyn Write> = match output {
        Some(path) => Box::new(File::create(path)?),
        None => Box::new(io::stdout()),
    };

    if pretty {
        serde_json::to_writer_pretty(&mut writer, value)?;
    } else {
        serde_json::to_writer(&mut writer, value)?;
    }

    writeln!(writer)?;
    Ok(())
}

pub fn write_text_output(value: &str, output: Option<PathBuf>) -> Result<()> {
    if let Some(path) = output {
        fs::write(&path, value)
            .with_context(|| format!("Failed to write to {}", path.display()))?;
    } else {
        let stdout = io::stdout();
        let mut handle = stdout.lock();
        handle.write_all(value.as_bytes())?;
        handle.write_all(b"\n")?;
    }
    Ok(())
}
