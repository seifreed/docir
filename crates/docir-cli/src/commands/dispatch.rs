use crate::{Cli, Commands};
use anyhow::Result;
use docir_app::ParserConfig;

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
            input,
            json,
            details,
            inventory,
            unknown,
            export,
            export_format,
            export_mode,
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
        } => run_query(
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            pretty,
            output,
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
        } => run_select(
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
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
        _ => unreachable!("command should be routed by run_command"),
    }
}

#[allow(clippy::too_many_arguments)]
fn run_coverage(
    input: std::path::PathBuf,
    json: bool,
    details: bool,
    inventory: bool,
    unknown: bool,
    export: Option<std::path::PathBuf>,
    export_format: crate::CoverageExportFormat,
    export_mode: crate::CoverageExportMode,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::coverage::run(
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
    )
}

#[allow(clippy::too_many_arguments)]
fn run_query(
    input: std::path::PathBuf,
    node_type: Option<String>,
    contains: Option<String>,
    format: Option<String>,
    has_external_refs: Option<bool>,
    has_macros: Option<bool>,
    pretty: bool,
    output: Option<std::path::PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::query::run(
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

#[allow(clippy::too_many_arguments)]
fn run_select(
    input: std::path::PathBuf,
    node_type: Option<String>,
    contains: Option<String>,
    format: Option<String>,
    has_external_refs: Option<bool>,
    has_macros: Option<bool>,
    pretty: bool,
    output: Option<std::path::PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::select::run(
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
