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
use std::path::{Path, PathBuf};

/// Public API entrypoint: parse_node_type.
pub fn parse_node_type(input: &str) -> Result<NodeType> {
    parse_core_node_type(input).map_err(|e| anyhow!(e))
}

/// Public API entrypoint: parse_doc_format.
pub fn parse_doc_format(input: &str) -> Result<DocumentFormat> {
    parse_core_document_format(input).map_err(|e| anyhow!(e))
}

/// Public API entrypoint: parse_node_id.
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

/// Public API entrypoint: build_app.
pub fn build_app(config: &ParserConfig) -> DocirApp {
    DocirApp::new(config.clone())
}

/// Public API entrypoint: parse_document.
pub fn parse_document(input: &PathBuf, parser_config: &ParserConfig) -> Result<ParsedDocument> {
    let (_, parsed) = build_app_and_parse(input, parser_config)?;
    Ok(parsed)
}

/// Public API entrypoint: build_app_and_parse.
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

/// Public API entrypoint: write_json_output.
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

pub(crate) fn run_json_document_command<T, F>(
    input: PathBuf,
    parser_config: &ParserConfig,
    pretty: bool,
    output: Option<PathBuf>,
    build: F,
) -> Result<()>
where
    T: Serialize,
    F: FnOnce(&ParsedDocument) -> Result<T>,
{
    let parsed = parse_document(&input, parser_config)?;
    let payload = build(&parsed)?;
    write_json_output(&payload, pretty, output)
}

pub(crate) fn run_json_app_command<T, F>(
    parser_config: &ParserConfig,
    pretty: bool,
    output: Option<PathBuf>,
    build: F,
) -> Result<()>
where
    T: Serialize,
    F: FnOnce(&DocirApp) -> Result<T>,
{
    let app = build_app(parser_config);
    let payload = build(&app)?;
    write_json_output(&payload, pretty, output)
}

pub(crate) fn run_text_document_command<T, F>(
    input: PathBuf,
    parser_config: &ParserConfig,
    output: Option<PathBuf>,
    build: F,
) -> Result<()>
where
    T: AsRef<str>,
    F: FnOnce(&DocirApp, &ParsedDocument) -> Result<T>,
{
    let (app, parsed) = build_app_and_parse(&input, parser_config)?;
    let output_data = build(&app, &parsed)?;
    write_text_output(output_data.as_ref(), output)
}

/// Public API entrypoint: write_text_output.
pub fn write_text_output(value: &str, output: Option<PathBuf>) -> Result<()> {
    let content = if value.ends_with('\n') {
        value.to_string()
    } else {
        format!("{}\n", value)
    };
    if let Some(path) = output {
        fs::write(&path, &content)
            .with_context(|| format!("Failed to write to {}", path.display()))?;
    } else {
        let stdout = io::stdout();
        let mut handle = stdout.lock();
        handle.write_all(content.as_bytes())?;
    }
    Ok(())
}

pub(crate) fn push_labeled_line(
    out: &mut String,
    indent: usize,
    label: &str,
    value: impl std::fmt::Display,
) {
    out.push_str(&" ".repeat(indent));
    out.push_str(label);
    out.push_str(": ");
    out.push_str(&value.to_string());
    out.push('\n');
}

pub(crate) fn push_bullet_line(
    out: &mut String,
    indent: usize,
    label: &str,
    value: impl std::fmt::Display,
) {
    out.push_str(&" ".repeat(indent));
    out.push_str("- ");
    out.push_str(label);
    out.push_str(": ");
    out.push_str(&value.to_string());
    out.push('\n');
}

pub(crate) fn source_format_label(input: &Path, fallback: &str) -> String {
    input
        .extension()
        .and_then(|value| value.to_str())
        .map(|value| value.to_ascii_lowercase())
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| fallback.to_string())
}

#[cfg(test)]
mod tests {
    use super::{
        parse_node_id, push_bullet_line, push_labeled_line, write_json_output, write_text_output,
    };
    use serde::Serialize;
    use std::fs;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn temp_file(name: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_{name}_{nanos}.tmp"))
    }

    #[test]
    fn parse_node_id_supports_raw_and_prefixed_values() {
        let raw = parse_node_id("42").expect("raw node id");
        let prefixed = parse_node_id("node_2a").expect("prefixed node id");
        assert_eq!(raw.as_u64(), 42);
        assert_eq!(prefixed.as_u64(), 42);
    }

    #[test]
    fn parse_node_id_rejects_invalid_input() {
        let err = parse_node_id("node_not_hex").expect_err("invalid id should fail");
        assert!(err.to_string().contains("Invalid node id"));
    }

    #[test]
    fn write_text_output_writes_to_file() {
        let path = temp_file("text_output");
        write_text_output("hello", Some(path.clone())).expect("write output");
        let written = fs::read_to_string(&path).expect("read output");
        assert_eq!(written, "hello\n");
        let _ = fs::remove_file(path);
    }

    #[derive(Serialize)]
    struct JsonProbe {
        ok: bool,
    }

    #[test]
    fn write_json_output_writes_json_document() {
        let path = temp_file("json_output");
        write_json_output(&JsonProbe { ok: true }, true, Some(path.clone())).expect("json write");
        let written = fs::read_to_string(&path).expect("read json");
        assert!(written.contains("\"ok\": true"));
        let _ = fs::remove_file(path);
    }

    #[test]
    fn push_labeled_and_bullet_lines_render_expected_text() {
        let mut out = String::new();
        push_labeled_line(&mut out, 2, "Path", "VBA/PROJECT");
        push_bullet_line(&mut out, 4, "CFB Stream", "VBA/PROJECT");
        assert_eq!(out, "  Path: VBA/PROJECT\n    - CFB Stream: VBA/PROJECT\n");
    }
}
