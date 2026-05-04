use crate::{Cli, Commands};
use anyhow::Result;
use docir_app::ParserConfig;

pub(crate) fn run(cli: Cli, parser_config: &ParserConfig) -> Result<()> {
    dispatch(cli.command, parser_config)
}

/// Routes CLI commands to their handler functions.
/// Each arm is a thin delegation with no logic beyond argument unpacking.
// NOTE: This function exceeds 80 LOC because it is a pure dispatch router
// with one arm per Commands variant. Decomposition would add indirection
// without reducing total LOC or improving readability.
fn dispatch(command: Commands, cfg: &ParserConfig) -> Result<()> {
    match command {
        // Core parse/summary/coverage
        Commands::Parse {
            input,
            format,
            output_opts,
        } => super::parse::run(input, format, output_opts, cfg),
        Commands::Summary { input } => super::summary::run(input, cfg),
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
        ),
        // Inspect
        Commands::Inventory { input, output_opts } => {
            super::inventory::run(input, output_opts, cfg)
        }
        Commands::ProbeFormat { input, output_opts } => {
            super::probe_format::run(input, output_opts, cfg)
        }
        Commands::ListTimes { input, output_opts } => {
            super::list_times::run(input, output_opts, cfg)
        }
        Commands::InspectMetadata { input, output_opts } => {
            super::inspect_metadata::run(input, output_opts, cfg)
        }
        Commands::InspectSheetRecords { input, output_opts } => {
            super::inspect_sheet_records::run(input, output_opts, cfg)
        }
        Commands::InspectSlideRecords { input, output_opts } => {
            super::inspect_slide_records::run(input, output_opts, cfg)
        }
        Commands::InspectDirectory { input, output_opts } => {
            super::inspect_directory::run(input, output_opts, cfg)
        }
        Commands::InspectSectors { input, output_opts } => {
            super::inspect_sectors::run(input, output_opts, cfg)
        }
        Commands::ReportIndicators { input, output_opts } => {
            super::report_indicators::run(input, output_opts, cfg)
        }
        // Extract
        Commands::ExtractLinks { input, output_opts } => {
            super::extract_links::run(input, output_opts, cfg)
        }
        Commands::ExtractFlash {
            input,
            out,
            overwrite,
            output_opts,
        } => super::extract_flash::run(input, out, overwrite, output_opts, cfg),
        Commands::Manifest { input, output_opts } => super::manifest::run(input, output_opts, cfg),
        Commands::DumpContainer { input, output_opts } => {
            super::dump_container::run(input, output_opts, cfg)
        }
        Commands::RecognizeVba {
            input,
            include_source,
            output_opts,
        } => super::recognize_vba::run(input, include_source, output_opts, cfg),
        Commands::ExtractVba {
            input,
            out,
            overwrite,
            best_effort,
        } => super::extract_vba::run(input, out, overwrite, best_effort, cfg),
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
        ),
        // Analysis
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
    }
}

#[cfg(test)]
mod tests {
    use super::dispatch;
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
    fn dispatch_routes_parse_summary_probe_inventory() {
        let config = ParserConfig::default();
        let parse_out = test_support::temp_file("parse", "json");
        dispatch(
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

        dispatch(
            Commands::Summary {
                input: test_support::fixture("minimal.docx"),
            },
            &config,
        )
        .expect("summary arm");

        dispatch(
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

        dispatch(
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

        let _ = fs::remove_file(parse_out);
    }

    #[test]
    fn dispatch_routes_inspect_commands() {
        let config = ParserConfig::default();
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

        dispatch(
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

        dispatch(
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

        dispatch(
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

        dispatch(
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

        let _ = fs::remove_file(list_times_input);
        let _ = fs::remove_file(metadata_input);
        let _ = fs::remove_file(inspect_directory_input);
        let _ = fs::remove_file(inspect_sectors_input);
    }

    #[test]
    fn dispatch_routes_output_commands() {
        let config = ParserConfig::default();

        dispatch(
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

        dispatch(
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

        dispatch(
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

        dispatch(
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

        dispatch(
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

        dispatch(
            Commands::Security {
                input: test_support::fixture("minimal.docx"),
                json: true,
                verbose: false,
            },
            &config,
        )
        .expect("security arm");
    }

    #[test]
    fn dispatch_routes_rules_diff_coverage() {
        let config = ParserConfig::default();
        let rules_out = test_support::temp_file("rules", "json");
        dispatch(
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
        dispatch(
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

        dispatch(
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

        let _ = fs::remove_file(rules_out);
        let _ = fs::remove_file(diff_out);
    }

    #[test]
    fn dispatch_routes_extract_artifacts() {
        let config = ParserConfig::default();
        let artifacts_out = test_support::temp_dir("extract_artifacts");
        dispatch(
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

        let _ = fs::remove_dir_all(artifacts_out);
    }

    #[test]
    fn dispatch_routes_query_select_grep_extract() {
        let config = ParserConfig::default();
        let query_out = test_support::temp_file("query", "json");
        dispatch(
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
        dispatch(
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
        dispatch(
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
        dispatch(
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
    fn dispatch_dump_node_invalid_id_fails() {
        let config = ParserConfig::default();
        let err = dispatch(
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
