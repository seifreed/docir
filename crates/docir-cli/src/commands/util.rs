//! Shared CLI helpers.

use anyhow::{anyhow, Result};
use docir_app::DocirApp;
use docir_app::ParserConfig;
use docir_core::types::{DocumentFormat, NodeType};
use serde::Serialize;
use std::fs::File;
use std::io::{self, Write};
use std::path::PathBuf;
use std::str::FromStr;

pub fn parse_node_type(input: &str) -> Result<NodeType> {
    NodeType::from_str(input).map_err(|e| anyhow!(e))
}

pub fn parse_doc_format(input: &str) -> Result<DocumentFormat> {
    DocumentFormat::from_str(input).map_err(|e| anyhow!(e))
}

pub fn build_app(config: &ParserConfig) -> DocirApp {
    DocirApp::new(config.clone())
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
