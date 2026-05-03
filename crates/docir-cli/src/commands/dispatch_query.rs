use anyhow::Result;
use docir_app::ParserConfig;
use std::path::PathBuf;

use super::query::QueryFilters;

pub(crate) fn cmd_diff(
    left: PathBuf,
    right: PathBuf,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::diff::run(left, right, pretty, output, parser_config)
}

pub(crate) fn cmd_rules(
    input: PathBuf,
    pretty: bool,
    output: Option<PathBuf>,
    profile: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::rules::run(input, pretty, output, profile, parser_config)
}

pub(crate) fn cmd_query(
    input: PathBuf,
    filters: QueryFilters,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::query::run_with_filters(input, filters, pretty, output, parser_config)
}

pub(crate) fn cmd_grep(
    input: PathBuf,
    pattern: String,
    node_type: Option<String>,
    format: Option<String>,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::grep::run(
        input,
        pattern,
        node_type,
        format,
        pretty,
        output,
        parser_config,
    )
}
