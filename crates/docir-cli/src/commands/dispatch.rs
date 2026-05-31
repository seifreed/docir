use crate::{Cli, Commands};
use anyhow::Result;
use docir_app::ParserConfig;

pub(crate) fn run(cli: Cli, parser_config: &ParserConfig) -> Result<()> {
    dispatch(cli.command, parser_config)
}

/// Routes CLI commands to their handler functions.
/// Each arm is a thin delegation with no logic beyond argument unpacking.
fn dispatch(command: Commands, cfg: &ParserConfig) -> Result<()> {
    let command = match dispatch_core_commands(command, cfg)? {
        Some(command) => command,
        None => return Ok(()),
    };
    let command = match dispatch_inspection_commands(command, cfg)? {
        Some(command) => command,
        None => return Ok(()),
    };
    let command = match dispatch_extraction_commands(command, cfg)? {
        Some(command) => command,
        None => return Ok(()),
    };
    dispatch_analysis_commands(command, cfg)
}

fn dispatch_core_commands(command: Commands, cfg: &ParserConfig) -> Result<Option<Commands>> {
    match command {
        Commands::Parse {
            input,
            format,
            output_opts,
        } => super::parse::run(input, format, output_opts, cfg).map(|()| None),
        Commands::Summary { input } => super::summary::run(input, cfg).map(|()| None),
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
            cfg,
        )
        .map(|()| None),
        command => Ok(Some(command)),
    }
}

fn dispatch_inspection_commands(command: Commands, cfg: &ParserConfig) -> Result<Option<Commands>> {
    match command {
        Commands::Inventory { input, output_opts } => {
            super::inventory::run(input, output_opts, cfg).map(|()| None)
        }
        Commands::ProbeFormat { input, output_opts } => {
            super::probe_format::run(input, output_opts, cfg).map(|()| None)
        }
        Commands::ListTimes { input, output_opts } => {
            super::list_times::run(input, output_opts, cfg).map(|()| None)
        }
        Commands::InspectMetadata { input, output_opts } => {
            super::inspect_metadata::run(input, output_opts, cfg).map(|()| None)
        }
        Commands::InspectSheetRecords { input, output_opts } => {
            super::inspect_sheet_records::run(input, output_opts, cfg).map(|()| None)
        }
        Commands::InspectSlideRecords { input, output_opts } => {
            super::inspect_slide_records::run(input, output_opts, cfg).map(|()| None)
        }
        Commands::InspectDirectory { input, output_opts } => {
            super::inspect_directory::run(input, output_opts, cfg).map(|()| None)
        }
        Commands::InspectSectors { input, output_opts } => {
            super::inspect_sectors::run(input, output_opts, cfg).map(|()| None)
        }
        Commands::ReportIndicators { input, output_opts } => {
            super::report_indicators::run(input, output_opts, cfg).map(|()| None)
        }
        command => Ok(Some(command)),
    }
}

fn dispatch_extraction_commands(command: Commands, cfg: &ParserConfig) -> Result<Option<Commands>> {
    match command {
        Commands::ExtractLinks { input, output_opts } => {
            super::extract_links::run(input, output_opts, cfg).map(|()| None)
        }
        Commands::ExtractFlash {
            input,
            out,
            overwrite,
            output_opts,
        } => super::extract_flash::run(input, out, overwrite, output_opts, cfg).map(|()| None),
        Commands::Manifest { input, output_opts } => {
            super::manifest::run(input, output_opts, cfg).map(|()| None)
        }
        Commands::DumpContainer { input, output_opts } => {
            super::dump_container::run(input, output_opts, cfg).map(|()| None)
        }
        Commands::RecognizeVba {
            input,
            include_source,
            output_opts,
        } => super::recognize_vba::run(input, include_source, output_opts, cfg).map(|()| None),
        Commands::ExtractVba {
            input,
            out,
            overwrite,
            best_effort,
        } => super::extract_vba::run(input, out, overwrite, best_effort, cfg).map(|()| None),
        Commands::ExtractArtifacts {
            input,
            out,
            overwrite,
            with_raw,
            no_media,
            only_ole,
            only_rtf_objects,
        } => super::extract_artifacts::run(
            input,
            out,
            super::extract_artifacts::ExtractArtifactsOptions {
                overwrite,
                with_raw,
                no_media,
                only_ole,
                only_rtf_objects,
            },
            cfg,
        )
        .map(|()| None),
        command => Ok(Some(command)),
    }
}

fn dispatch_analysis_commands(command: Commands, cfg: &ParserConfig) -> Result<()> {
    match command {
        Commands::Security {
            input,
            json,
            verbose,
        } => super::security::run(input, json, verbose, cfg),
        Commands::DumpNode {
            input,
            node_id,
            format,
        } => super::dump_node::run(input, &node_id, format, cfg),
        Commands::Diff {
            left,
            right,
            output_opts,
        } => super::diff::run(left, right, output_opts, cfg),
        Commands::Rules {
            input,
            output_opts,
            profile,
        } => super::rules::run(input, output_opts, profile, cfg),
        Commands::Query {
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            output_opts,
        }
        | Commands::Select {
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            output_opts,
        } => super::query::run_with_filters(
            input,
            super::query::QueryFilters {
                node_type,
                contains,
                format,
                has_external_refs,
                has_macros,
            },
            output_opts,
            cfg,
        ),
        Commands::Grep {
            input,
            pattern,
            node_type,
            format,
            output_opts,
        } => super::grep::run(input, pattern, node_type, format, output_opts, cfg),
        Commands::Extract {
            input,
            node_id,
            node_type,
            output_opts,
        } => super::extract::run(input, node_id, node_type, output_opts, cfg),
        _command => anyhow::bail!("internal CLI routing error"),
    }
}

#[cfg(test)]
#[path = "dispatch_tests.rs"]
mod tests;
