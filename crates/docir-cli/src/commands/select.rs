//! Select nodes from IR using query predicates.

use anyhow::Result;
use docir_app::ParserConfig;
use std::path::PathBuf;

use crate::commands::query::{run_with_filters, QueryFilters};

pub fn run(
    input: PathBuf,
    node_type: Option<String>,
    contains: Option<String>,
    format: Option<String>,
    has_external_refs: Option<bool>,
    has_macros: Option<bool>,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    run_with_filters(
        input,
        QueryFilters {
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
        },
        pretty,
        output,
        parser_config,
    )
}
