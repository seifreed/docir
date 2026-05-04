//! Parse command implementation.

use anyhow::{Context, Result};
use docir_app::ParserConfig;
use std::path::PathBuf;

use crate::cli::PrettyOutputOpts;
use crate::commands::util::run_text_document_command;
use crate::OutputFormat;

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    format: OutputFormat,
    opts: PrettyOutputOpts,
    parser_config: &ParserConfig,
) -> Result<()> {
    let PrettyOutputOpts { pretty, output } = opts;
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
    use crate::cli::PrettyOutputOpts;
    use crate::test_support;
    use crate::OutputFormat;
    use docir_app::ParserConfig;
    use std::fs;

    #[test]
    fn parse_run_writes_json_output_file() {
        let input = test_support::fixture("minimal.docx");
        let output = test_support::temp_file("minimal", "json");
        run(
            input,
            OutputFormat::Json,
            PrettyOutputOpts {
                pretty: true,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("parse command should succeed");
        let text = fs::read_to_string(&output).expect("output file");
        assert!(text.contains('{'));
        let _ = fs::remove_file(output);
    }
}
