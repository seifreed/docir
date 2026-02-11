//! Diff two documents and output the IR diff.

use crate::commands::util::{build_parser, write_json_output};
use anyhow::Result;
use docir_diff::DiffEngine;
use docir_parser::ParserConfig;
use std::path::PathBuf;

pub fn run(
    left: PathBuf,
    right: PathBuf,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let parser = build_parser(parser_config);

    let left_doc = parser.parse_file(&left)?;
    let right_doc = parser.parse_file(&right)?;

    let diff = DiffEngine::diff(
        &left_doc.store,
        left_doc.root_id,
        &right_doc.store,
        right_doc.root_id,
    );

    write_json_output(&diff, pretty, output)
}
