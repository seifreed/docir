//! Diff two documents and output the IR diff.

use crate::commands::util::run_json_app_command;
use anyhow::Result;
use docir_app::ParserConfig;
use std::path::PathBuf;

/// Public API entrypoint: run.
pub fn run(
    left: PathBuf,
    right: PathBuf,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    run_json_app_command(parser_config, pretty, output, move |app| {
        let left_doc = app.parse_file(&left)?;
        let right_doc = app.parse_file(&right)?;
        Ok(app.diff(&left_doc, &right_doc)?)
    })
}
