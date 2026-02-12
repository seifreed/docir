//! Diff two documents and output the IR diff.

use crate::commands::util::{build_app, write_json_output};
use anyhow::Result;
use docir_app::ParserConfig;
use std::path::PathBuf;

pub fn run(
    left: PathBuf,
    right: PathBuf,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let app = build_app(parser_config);

    let left_doc = app.parse_file(&left)?;
    let right_doc = app.parse_file(&right)?;

    let diff = app.diff(&left_doc, &right_doc);

    write_json_output(&diff, pretty, output)
}
