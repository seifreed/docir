use crate::{Cli, Commands};
use anyhow::Result;
use docir_app::ParserConfig;
use std::path::PathBuf;

struct CoverageCommand {
    input: PathBuf,
    options: super::coverage::CoverageOptions,
}

struct QueryLikeCommand {
    input: PathBuf,
    filters: super::query::QueryFilters,
    pretty: bool,
    output: Option<PathBuf>,
}

pub(crate) fn run(cli: Cli, parser_config: &ParserConfig) -> Result<()> {
    run_command(cli.command, parser_config)
}

fn run_command(command: Commands, parser_config: &ParserConfig) -> Result<()> {
    match command {
        Commands::Parse {
            input,
            format,
            pretty,
            output,
        } => super::parse::run(input, format, pretty, output, parser_config),
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
        } => run_coverage(
            CoverageCommand {
                input,
                options: super::coverage::CoverageOptions {
                    json,
                    details,
                    inventory,
                    unknown,
                    export,
                    export_format,
                    export_mode,
                },
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
        other => run_query_extract_commands(other, parser_config),
    }
}

fn run_query_extract_commands(command: Commands, parser_config: &ParserConfig) -> Result<()> {
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
        } => run_query_like(
            QueryLikeCommand {
                input,
                filters: super::query::QueryFilters {
                    node_type,
                    contains,
                    format,
                    has_external_refs,
                    has_macros,
                },
                pretty,
                output,
            },
            parser_config,
        ),
        Commands::Select {
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            pretty,
            output,
        } => run_query_like(
            QueryLikeCommand {
                input,
                filters: super::query::QueryFilters {
                    node_type,
                    contains,
                    format,
                    has_external_refs,
                    has_macros,
                },
                pretty,
                output,
            },
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
        _ => unreachable!("command should be routed by run_command"),
    }
}

fn run_coverage(command: CoverageCommand, parser_config: &ParserConfig) -> Result<()> {
    super::coverage::run(command.input, command.options, parser_config)
}

fn run_query_like(command: QueryLikeCommand, parser_config: &ParserConfig) -> Result<()> {
    super::query::run_with_filters(
        command.input,
        command.filters,
        command.pretty,
        command.output,
        parser_config,
    )
}
