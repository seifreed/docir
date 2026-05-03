use super::dispatch_extract::{
    cmd_extract, cmd_extract_artifacts, cmd_extract_flash, cmd_extract_links, cmd_extract_vba,
};
use super::dispatch_inspect::{
    cmd_inspect_directory, cmd_inspect_metadata, cmd_inspect_sectors, cmd_inspect_sheet_records,
    cmd_inspect_slide_records,
};
use crate::{Cli, Commands};
use anyhow::Result;
use docir_app::ParserConfig;
use std::path::PathBuf;

pub(crate) fn run(cli: Cli, parser_config: &ParserConfig) -> Result<()> {
    run_command(cli.command, parser_config)
}

fn run_command(command: Commands, parser_config: &ParserConfig) -> Result<()> {
    match command {
        Commands::Parse {
            input,
            format,
            output_opts,
        } => cmd_parse(
            input,
            format,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::Summary { input } => cmd_summary(input, parser_config),
        Commands::Coverage {
            input,
            json,
            details,
            inventory,
            unknown,
            export,
            export_format,
            export_mode,
        } => cmd_coverage(
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
        Commands::Inventory { input, output_opts } => cmd_inventory(
            input,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::ProbeFormat { input, output_opts } => cmd_probe_format(
            input,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::ListTimes { input, output_opts } => cmd_list_times(
            input,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::InspectMetadata { input, output_opts } => cmd_inspect_metadata(
            input,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::InspectSheetRecords { input, output_opts } => cmd_inspect_sheet_records(
            input,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::InspectSlideRecords { input, output_opts } => cmd_inspect_slide_records(
            input,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::InspectDirectory { input, output_opts } => cmd_inspect_directory(
            input,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::InspectSectors { input, output_opts } => cmd_inspect_sectors(
            input,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::ReportIndicators { input, output_opts } => cmd_report_indicators(
            input,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::ExtractLinks { input, output_opts } => cmd_extract_links(
            input,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::ExtractFlash {
            input,
            out,
            overwrite,
            output_opts,
        } => cmd_extract_flash(
            input,
            out,
            overwrite,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::Manifest { input, output_opts } => {
            cmd_manifest(input, output_opts.pretty, output_opts.output, parser_config)
        }
        Commands::DumpContainer { input, output_opts } => cmd_dump_container(
            input,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::RecognizeVba {
            input,
            include_source,
            output_opts,
        } => cmd_recognize_vba(
            input,
            include_source,
            output_opts.json,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::ExtractVba {
            input,
            out,
            overwrite,
            best_effort,
        } => cmd_extract_vba(input, out, overwrite, best_effort, parser_config),
        Commands::ExtractArtifacts {
            input,
            out,
            overwrite,
            with_raw,
            no_media,
            only_ole,
            only_rtf_objects,
        } => cmd_extract_artifacts(
            input,
            out,
            overwrite,
            with_raw,
            no_media,
            only_ole,
            only_rtf_objects,
            parser_config,
        ),
        Commands::Security {
            input,
            json,
            verbose,
        } => cmd_security(input, json, verbose, parser_config),
        Commands::DumpNode {
            input,
            node_id,
            format,
        } => cmd_dump_node(input, node_id, format, parser_config),
        Commands::Diff {
            left,
            right,
            output_opts,
        } => cmd_diff(
            left,
            right,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::Rules {
            input,
            output_opts,
            profile,
        } => cmd_rules(
            input,
            output_opts.pretty,
            output_opts.output,
            profile,
            parser_config,
        ),
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
        } => cmd_query(
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::Grep {
            input,
            pattern,
            node_type,
            format,
            output_opts,
        } => cmd_grep(
            input,
            pattern,
            node_type,
            format,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
        Commands::Extract {
            input,
            node_id,
            node_type,
            output_opts,
        } => cmd_extract(
            input,
            node_id,
            node_type,
            output_opts.pretty,
            output_opts.output,
            parser_config,
        ),
    }
}

fn cmd_parse(
    input: PathBuf,
    format: crate::OutputFormat,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::parse::run(input, format, pretty, output, parser_config)
}

fn cmd_summary(input: PathBuf, parser_config: &ParserConfig) -> Result<()> {
    super::summary::run(input, parser_config)
}

#[allow(clippy::too_many_arguments)]
fn cmd_coverage(
    input: PathBuf,
    json: bool,
    details: bool,
    inventory: bool,
    unknown: bool,
    export: Option<PathBuf>,
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

fn cmd_inventory(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::inventory::run(input, json, pretty, output, parser_config)
}

fn cmd_probe_format(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::probe_format::run(input, json, pretty, output, parser_config)
}

fn cmd_list_times(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::list_times::run(input, json, pretty, output, parser_config)
}

fn cmd_report_indicators(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::report_indicators::run(input, json, pretty, output, parser_config)
}

fn cmd_manifest(
    input: PathBuf,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::manifest::run(input, pretty, output, parser_config)
}

fn cmd_dump_container(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::dump_container::run(input, json, pretty, output, parser_config)
}

fn cmd_recognize_vba(
    input: PathBuf,
    include_source: bool,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::recognize_vba::run(input, include_source, json, pretty, output, parser_config)
}

fn cmd_security(
    input: PathBuf,
    json: bool,
    verbose: bool,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::security::run(input, json, verbose, parser_config)
}

fn cmd_dump_node(
    input: PathBuf,
    node_id: String,
    format: crate::OutputFormat,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::dump_node::run(input, &node_id, format, parser_config)
}

fn cmd_diff(
    left: PathBuf,
    right: PathBuf,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::diff::run(left, right, pretty, output, parser_config)
}

fn cmd_rules(
    input: PathBuf,
    pretty: bool,
    output: Option<PathBuf>,
    profile: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::rules::run(input, pretty, output, profile, parser_config)
}

#[allow(clippy::too_many_arguments)]
fn cmd_query(
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
    super::query::run_with_filters(
        input,
        super::query::QueryFilters {
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

#[allow(clippy::too_many_arguments)]
fn cmd_grep(
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

#[cfg(test)]
mod tests {
    use super::run_command;
    use crate::test_support;
    use crate::{
        Commands, CoverageExportFormat, CoverageExportMode, JsonOutputOpts, OutputFormat,
        PrettyOutputOpts,
    };
    use docir_app::{
        test_support::{
            build_test_cfb, build_test_cfb_with_times, build_test_property_set_stream,
            TestPropertyValue,
        },
        ParserConfig,
    };
    use std::fs;

    #[test]
    fn run_command_routes_main_arms_successfully() {
        let config = ParserConfig::default();
        let parse_out = test_support::temp_file("parse", "json");
        let list_times_input = test_support::temp_file("times", "doc");
        let metadata_input = test_support::temp_file("metadata", "doc");
        let inspect_directory_input = test_support::temp_file("inspect_directory", "doc");
        let inspect_sectors_input = test_support::temp_file("inspect_sectors", "doc");
        fs::write(
            &list_times_input,
            build_test_cfb_with_times(&[("WordDocument", b"doc")], &[("WordDocument", 10, 20)]),
        )
        .expect("list-times fixture");
        fs::write(
            &metadata_input,
            build_test_cfb(&[(
                "\u{0005}SummaryInformation",
                &build_test_property_set_stream(&[(2, TestPropertyValue::Str("Specimen"))]),
            )]),
        )
        .expect("metadata fixture");
        fs::write(
            &inspect_directory_input,
            build_test_cfb(&[
                ("WordDocument", b"doc"),
                ("VBA/PROJECT", b"meta"),
                ("ObjectPool/1/Ole10Native", b"payload"),
            ]),
        )
        .expect("inspect-directory fixture");
        fs::write(
            &inspect_sectors_input,
            build_test_cfb(&[("WordDocument", b"doc"), ("VBA/PROJECT", b"meta")]),
        )
        .expect("inspect-sectors fixture");
        run_command(
            Commands::Parse {
                input: test_support::fixture("minimal.docx"),
                format: OutputFormat::Json,
                output_opts: PrettyOutputOpts {
                    pretty: true,
                    output: Some(parse_out.clone()),
                },
            },
            &config,
        )
        .expect("parse arm");
        assert!(parse_out.exists());

        run_command(
            Commands::Summary {
                input: test_support::fixture("minimal.docx"),
            },
            &config,
        )
        .expect("summary arm");

        run_command(
            Commands::ProbeFormat {
                input: test_support::fixture("minimal.docx"),
                output_opts: JsonOutputOpts {
                    json: true,
                    pretty: true,
                    output: None,
                },
            },
            &config,
        )
        .expect("probe-format arm");

        run_command(
            Commands::Inventory {
                input: test_support::fixture("minimal.docx"),
                output_opts: JsonOutputOpts {
                    json: true,
                    pretty: true,
                    output: None,
                },
            },
            &config,
        )
        .expect("inventory arm");

        run_command(
            Commands::ListTimes {
                input: list_times_input.clone(),
                output_opts: JsonOutputOpts {
                    json: true,
                    pretty: true,
                    output: None,
                },
            },
            &config,
        )
        .expect("list-times arm");

        run_command(
            Commands::Manifest {
                input: test_support::fixture("minimal.docx"),
                output_opts: PrettyOutputOpts {
                    pretty: true,
                    output: None,
                },
            },
            &config,
        )
        .expect("manifest arm");

        run_command(
            Commands::InspectMetadata {
                input: metadata_input.clone(),
                output_opts: JsonOutputOpts {
                    json: true,
                    pretty: true,
                    output: None,
                },
            },
            &config,
        )
        .expect("inspect-metadata arm");

        run_command(
            Commands::InspectDirectory {
                input: inspect_directory_input.clone(),
                output_opts: JsonOutputOpts {
                    json: true,
                    pretty: true,
                    output: None,
                },
            },
            &config,
        )
        .expect("inspect-directory arm");

        run_command(
            Commands::InspectSectors {
                input: inspect_sectors_input.clone(),
                output_opts: JsonOutputOpts {
                    json: true,
                    pretty: true,
                    output: None,
                },
            },
            &config,
        )
        .expect("inspect-sectors arm");

        run_command(
            Commands::ReportIndicators {
                input: test_support::fixture("minimal.docx"),
                output_opts: JsonOutputOpts {
                    json: true,
                    pretty: true,
                    output: None,
                },
            },
            &config,
        )
        .expect("report-indicators arm");

        run_command(
            Commands::ExtractLinks {
                input: test_support::fixture("minimal.docx"),
                output_opts: JsonOutputOpts {
                    json: true,
                    pretty: true,
                    output: None,
                },
            },
            &config,
        )
        .expect("extract-links arm");

        run_command(
            Commands::DumpContainer {
                input: test_support::fixture("minimal.docx"),
                output_opts: JsonOutputOpts {
                    json: true,
                    pretty: true,
                    output: None,
                },
            },
            &config,
        )
        .expect("dump-container arm");

        run_command(
            Commands::RecognizeVba {
                input: test_support::fixture("minimal.docx"),
                include_source: false,
                output_opts: JsonOutputOpts {
                    json: true,
                    pretty: true,
                    output: None,
                },
            },
            &config,
        )
        .expect("recognize-vba arm");

        run_command(
            Commands::Security {
                input: test_support::fixture("minimal.docx"),
                json: true,
                verbose: false,
            },
            &config,
        )
        .expect("security arm");

        let rules_out = test_support::temp_file("rules", "json");
        run_command(
            Commands::Rules {
                input: test_support::fixture("minimal.docx"),
                output_opts: PrettyOutputOpts {
                    pretty: true,
                    output: Some(rules_out.clone()),
                },
                profile: None,
            },
            &config,
        )
        .expect("rules arm");
        assert!(rules_out.exists());

        let diff_out = test_support::temp_file("diff", "json");
        run_command(
            Commands::Diff {
                left: test_support::fixture("minimal.docx"),
                right: test_support::fixture("minimal.docx"),
                output_opts: PrettyOutputOpts {
                    pretty: true,
                    output: Some(diff_out.clone()),
                },
            },
            &config,
        )
        .expect("diff arm");
        assert!(diff_out.exists());

        run_command(
            Commands::Coverage {
                input: test_support::fixture("minimal.docx"),
                json: true,
                details: false,
                inventory: false,
                unknown: false,
                export: None,
                export_format: CoverageExportFormat::Json,
                export_mode: CoverageExportMode::Full,
            },
            &config,
        )
        .expect("coverage arm");

        let artifacts_out = test_support::temp_dir("extract_artifacts");
        run_command(
            Commands::ExtractArtifacts {
                input: test_support::fixture("minimal.docx"),
                out: artifacts_out.clone(),
                overwrite: false,
                with_raw: false,
                no_media: false,
                only_ole: false,
                only_rtf_objects: false,
            },
            &config,
        )
        .expect("extract artifacts arm");
        assert!(artifacts_out.join("manifest.json").exists());

        let _ = fs::remove_file(parse_out);
        let _ = fs::remove_file(list_times_input);
        let _ = fs::remove_file(metadata_input);
        let _ = fs::remove_file(inspect_directory_input);
        let _ = fs::remove_file(inspect_sectors_input);
        let _ = fs::remove_file(rules_out);
        let _ = fs::remove_file(diff_out);
        let _ = fs::remove_dir_all(artifacts_out);
    }

    #[test]
    fn run_command_routes_query_extract_arms_successfully() {
        let config = ParserConfig::default();
        let query_out = test_support::temp_file("query", "json");
        run_command(
            Commands::Query {
                input: test_support::fixture("minimal.docx"),
                node_type: Some("Paragraph".to_string()),
                contains: Some("Hello".to_string()),
                format: Some("docx".to_string()),
                has_external_refs: None,
                has_macros: None,
                output_opts: PrettyOutputOpts {
                    pretty: true,
                    output: Some(query_out.clone()),
                },
            },
            &config,
        )
        .expect("query arm");
        assert!(query_out.exists());

        let select_out = test_support::temp_file("select", "json");
        run_command(
            Commands::Select {
                input: test_support::fixture("minimal.docx"),
                node_type: Some("Paragraph".to_string()),
                contains: None,
                format: None,
                has_external_refs: None,
                has_macros: None,
                output_opts: PrettyOutputOpts {
                    pretty: true,
                    output: Some(select_out.clone()),
                },
            },
            &config,
        )
        .expect("select arm");
        assert!(select_out.exists());

        let grep_out = test_support::temp_file("grep", "json");
        run_command(
            Commands::Grep {
                input: test_support::fixture("minimal.docx"),
                pattern: "Hello".to_string(),
                node_type: None,
                format: None,
                output_opts: PrettyOutputOpts {
                    pretty: true,
                    output: Some(grep_out.clone()),
                },
            },
            &config,
        )
        .expect("grep arm");
        assert!(grep_out.exists());

        let extract_out = test_support::temp_file("extract", "json");
        run_command(
            Commands::Extract {
                input: test_support::fixture("minimal.docx"),
                node_id: Vec::new(),
                node_type: Some("Paragraph".to_string()),
                output_opts: PrettyOutputOpts {
                    pretty: true,
                    output: Some(extract_out.clone()),
                },
            },
            &config,
        )
        .expect("extract arm");
        assert!(extract_out.exists());

        let _ = fs::remove_file(query_out);
        let _ = fs::remove_file(select_out);
        let _ = fs::remove_file(grep_out);
        let _ = fs::remove_file(extract_out);
    }

    #[test]
    fn run_command_dump_node_invalid_id_fails() {
        let config = ParserConfig::default();
        let err = run_command(
            Commands::DumpNode {
                input: test_support::fixture("minimal.docx"),
                node_id: "invalid".to_string(),
                format: OutputFormat::Json,
            },
            &config,
        )
        .expect_err("invalid dump-node id should fail");
        assert!(err.to_string().contains("Invalid node ID format"));
    }
}
