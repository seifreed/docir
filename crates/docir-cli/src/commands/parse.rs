//! Parse command implementation.

use anyhow::{Context, Result};
use docir_app::ParserConfig;
use std::path::PathBuf;

use crate::commands::util::run_text_document_command;
use crate::OutputFormat;

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    format: OutputFormat,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let output_path = output.as_ref().map(|path| path.display().to_string());
    run_text_document_command(
        input,
        parser_config,
        output,
        move |app, parsed| match format {
            OutputFormat::Json => app
                .serialize_json(parsed, pretty)
                .context("Failed to serialize to JSON"),
        },
    )?;
    if let Some(path) = output_path {
        eprintln!("Output written to {}", path);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::run;
    use crate::OutputFormat;
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
        std::env::temp_dir().join(format!("docir_cli_parse_{name}_{nanos}.json"))
    }

    #[test]
    fn parse_run_writes_json_output_file() {
        let input = fixture("minimal.docx");
        let output = temp_file("minimal");
        run(
            input,
            OutputFormat::Json,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("parse command should succeed");
        let text = fs::read_to_string(&output).expect("output file");
        assert!(text.contains('{'));
        let _ = fs::remove_file(output);
    }
}
