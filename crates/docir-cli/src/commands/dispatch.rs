use super::query::QueryFilters;
use crate::{Cli, Commands};
use anyhow::Result;
use docir_app::ParserConfig;
use std::path::PathBuf;

pub(crate) fn run(cli: Cli, parser_config: &ParserConfig) -> Result<()> {
    dispatch_command(cli.command, parser_config)
}

fn dispatch_command(command: Commands, parser_config: &ParserConfig) -> Result<()> {
    match command {
        Commands::Parse {
            input,
            format,
            pretty,
            output,
        } => super::parse::run(input, format, pretty, output, parser_config),
        other => dispatch_non_parse(other, parser_config),
    }
}

fn dispatch_non_parse(command: Commands, parser_config: &ParserConfig) -> Result<()> {
    match command {
        Commands::Summary { input } => super::summary::run(input, parser_config),
        Commands::Coverage {
            input,
            json,
            details,
            inventory,
            unknown,
            export,
            export_format,
            export_mode,
        } => super::coverage::run(
            input,
            super::coverage::CoverageOptions {
                json,
                details,
                inventory,
                unknown,
                export,
                export_format,
                export_mode,
            },
            parser_config,
        ),
        Commands::Security {
            input,
            json,
            verbose,
        } => super::security::run(input, json, verbose, parser_config),
        Commands::DumpNode {
            input,
            node_id,
            format,
        } => super::dump_node::run(input, &node_id, format, parser_config),
        Commands::Diff {
            left,
            right,
            pretty,
            output,
        } => super::diff::run(left, right, pretty, output, parser_config),
        Commands::Rules {
            input,
            pretty,
            output,
            profile,
        } => super::rules::run(input, pretty, output, profile, parser_config),
        other => dispatch_query_extract(other, parser_config),
    }
}

fn dispatch_query_extract(command: Commands, parser_config: &ParserConfig) -> Result<()> {
    match command {
        Commands::Query {
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            pretty,
            output,
        }
        | Commands::Select {
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            pretty,
            output,
        } => run_query_with_filters(
            input,
            query_filters(node_type, contains, format, has_external_refs, has_macros),
            pretty,
            output,
            parser_config,
        ),
        Commands::Grep {
            input,
            pattern,
            node_type,
            format,
            pretty,
            output,
        } => super::grep::run(
            input,
            pattern,
            node_type,
            format,
            pretty,
            output,
            parser_config,
        ),
        Commands::Extract {
            input,
            node_id,
            node_type,
            pretty,
            output,
        } => super::extract::run(input, node_id, node_type, pretty, output, parser_config),
        Commands::Parse { .. }
        | Commands::Summary { .. }
        | Commands::Coverage { .. }
        | Commands::Security { .. }
        | Commands::DumpNode { .. }
        | Commands::Diff { .. }
        | Commands::Rules { .. } => unreachable!("command handled in prior dispatcher"),
    }
}

fn run_query_with_filters(
    input: PathBuf,
    filters: QueryFilters,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::query::run_with_filters(input, filters, pretty, output, parser_config)
}

fn query_filters(
    node_type: Option<String>,
    contains: Option<String>,
    format: Option<String>,
    has_external_refs: Option<bool>,
    has_macros: Option<bool>,
) -> QueryFilters {
    QueryFilters {
        node_type,
        contains,
        format,
        has_external_refs,
        has_macros,
    }
}
