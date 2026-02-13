use super::query::QueryFilters;
use crate::{Cli, Commands};
use anyhow::Result;
use docir_app::ParserConfig;

pub(crate) fn run(cli: Cli, parser_config: &ParserConfig) -> Result<()> {
    match cli.command {
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

        Commands::Query {
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            pretty,
            output,
        } => super::query::run_with_filters(
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
        } => super::query::run_with_filters(
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
    }
}
