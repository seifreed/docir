//! Select nodes from IR using query predicates.

use anyhow::Result;
use docir_app::ParserConfig;
use std::path::PathBuf;

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
    crate::commands::query::run(
        input,
        node_type,
        contains,
        format,
        has_external_refs,
        has_macros,
        pretty,
        output,
        parser_config,
    )
}
