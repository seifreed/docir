//! # docir CLI
//!
//! Command-line interface for the docir document analysis toolkit.

mod cli;
mod commands;

use anyhow::Result;
use clap::Parser;

pub(crate) use cli::{
    build_parser_config, Cli, Commands, CoverageExportFormat, CoverageExportMode, OutputFormat,
};

fn main() -> Result<()> {
    env_logger::Builder::from_env(env_logger::Env::default().default_filter_or("warn")).init();

    let cli = Cli::parse();
    let parser_config = build_parser_config(&cli);

    match cli.command {
        Commands::Parse {
            input,
            format,
            pretty,
            output,
        } => commands::parse::run(input, format, pretty, output, &parser_config),

        Commands::Summary { input } => commands::summary::run(input, &parser_config),
        Commands::Coverage {
            input,
            json,
            details,
            inventory,
            unknown,
            export,
            export_format,
            export_mode,
        } => commands::coverage::run(
            input,
            commands::coverage::CoverageOptions {
                json,
                details,
                inventory,
                unknown,
                export,
                export_format,
                export_mode,
            },
            &parser_config,
        ),

        Commands::Security {
            input,
            json,
            verbose,
        } => commands::security::run(input, json, verbose, &parser_config),

        Commands::DumpNode {
            input,
            node_id,
            format,
        } => commands::dump_node::run(input, &node_id, format, &parser_config),

        Commands::Diff {
            left,
            right,
            pretty,
            output,
        } => commands::diff::run(left, right, pretty, output, &parser_config),

        Commands::Rules {
            input,
            pretty,
            output,
            profile,
        } => commands::rules::run(input, pretty, output, profile, &parser_config),

        Commands::Query {
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            pretty,
            output,
        } => commands::query::run(
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            pretty,
            output,
            &parser_config,
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
        } => commands::select::run(
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            pretty,
            output,
            &parser_config,
        ),

        Commands::Grep {
            input,
            pattern,
            node_type,
            format,
            pretty,
            output,
        } => commands::grep::run(
            input,
            pattern,
            node_type,
            format,
            pretty,
            output,
            &parser_config,
        ),

        Commands::Extract {
            input,
            node_id,
            node_type,
            pretty,
            output,
        } => commands::extract::run(input, node_id, node_type, pretty, output, &parser_config),
    }
}
