//! Inspect CFB sector allocation and stream chains.

use anyhow::Result;
use docir_app::{ParserConfig, inspect_sectors_path};
use std::path::PathBuf;

use crate::cli::JsonOutputOpts;
use crate::commands::util::run_app_path_dual_output;

#[path = "inspect_sectors_format.rs"]
mod inspect_sectors_format;
use inspect_sectors_format::format_inspection_text;

#[cfg(test)]
#[path = "inspect_sectors_tests.rs"]
mod tests;

/// Public API entrypoint: run.
pub fn run(input: PathBuf, opts: JsonOutputOpts, parser_config: &ParserConfig) -> Result<()> {
    run_app_path_dual_output(
        input,
        opts,
        parser_config,
        "inspection",
        inspect_sectors_path,
        format_inspection_text,
    )
}
