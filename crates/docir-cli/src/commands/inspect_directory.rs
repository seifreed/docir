//! Inspect normal CFB directory entries.

use anyhow::Result;
use docir_app::{ParserConfig, inspect_directory_path};
use std::path::PathBuf;

use crate::cli::JsonOutputOpts;
use crate::commands::util::run_app_path_dual_output;

#[path = "inspect_directory_format.rs"]
mod inspect_directory_format;
use inspect_directory_format::format_inspection_text;

#[cfg(test)]
#[path = "inspect_directory_tests.rs"]
mod tests;

/// Public API entrypoint: run.
pub fn run(input: PathBuf, opts: JsonOutputOpts, parser_config: &ParserConfig) -> Result<()> {
    run_app_path_dual_output(
        input,
        opts,
        parser_config,
        "inspection",
        inspect_directory_path,
        format_inspection_text,
    )
}
