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
        Commands::Inventory {
            input,
            json,
            pretty,
            output,
        } => super::inventory::run(input, json, pretty, output, parser_config),
        Commands::ProbeFormat {
            input,
            json,
            pretty,
            output,
        } => super::probe_format::run(input, json, pretty, output, parser_config),
        Commands::ListTimes {
            input,
            json,
            pretty,
            output,
        } => super::list_times::run(input, json, pretty, output, parser_config),
        Commands::InspectMetadata {
            input,
            json,
            pretty,
            output,
        } => super::inspect_metadata::run(input, json, pretty, output, parser_config),
        Commands::InspectSheetRecords {
            input,
            json,
            pretty,
            output,
        } => super::inspect_sheet_records::run(input, json, pretty, output, parser_config),
        Commands::InspectSlideRecords {
            input,
            json,
            pretty,
            output,
        } => super::inspect_slide_records::run(input, json, pretty, output, parser_config),
        Commands::InspectDirectory {
            input,
            json,
            pretty,
            output,
        } => super::inspect_directory::run(input, json, pretty, output, parser_config),
        Commands::InspectSectors {
            input,
            json,
            pretty,
            output,
        } => super::inspect_sectors::run(input, json, pretty, output, parser_config),
        Commands::ReportIndicators {
            input,
            json,
            pretty,
            output,
        } => super::report_indicators::run(input, json, pretty, output, parser_config),
        Commands::ExtractLinks {
            input,
            json,
            pretty,
            output,
        } => super::extract_links::run(input, json, pretty, output, parser_config),
        Commands::ExtractFlash {
            input,
            out,
            overwrite,
            json,
            pretty,
            output,
        } => super::extract_flash::run(input, out, overwrite, json, pretty, output, parser_config),
        Commands::Manifest {
            input,
            pretty,
            output,
        } => super::manifest::run(input, pretty, output, parser_config),
        Commands::DumpContainer {
            input,
            json,
            pretty,
            output,
        } => super::dump_container::run(input, json, pretty, output, parser_config),
        Commands::RecognizeVba {
            input,
            include_source,
            json,
            pretty,
            output,
        } => super::recognize_vba::run(input, include_source, json, pretty, output, parser_config),
        Commands::ExtractVba {
            input,
            out,
            overwrite,
            best_effort,
        } => super::extract_vba::run(input, out, overwrite, best_effort, parser_config),
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
        } => super::query::run_with_filters(
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

#[cfg(test)]
mod tests {
    use super::run_command;
    use crate::{Commands, CoverageExportFormat, CoverageExportMode, OutputFormat};
    use docir_app::{
        test_support::{
            build_test_cfb, build_test_cfb_with_times, build_test_property_set_stream,
            TestPropertyValue,
        },
        ParserConfig,
    };
    use std::fs;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn fixture(name: &str) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../fixtures/ooxml")
            .join(name)
    }

    fn temp_file(name: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_dispatch_{name}_{nanos}.json"))
    }

    fn temp_dir(name: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_dispatch_{name}_{nanos}"))
    }

    fn temp_input(name: &str, ext: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_dispatch_{name}_{nanos}.{ext}"))
    }

    #[test]
    fn run_command_routes_main_arms_successfully() {
        let config = ParserConfig::default();
        let parse_out = temp_file("parse");
        let list_times_input = temp_input("times", "doc");
        let metadata_input = temp_input("metadata", "doc");
        let inspect_directory_input = temp_input("inspect_directory", "doc");
        let inspect_sectors_input = temp_input("inspect_sectors", "doc");
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
                input: fixture("minimal.docx"),
                format: OutputFormat::Json,
                pretty: true,
                output: Some(parse_out.clone()),
            },
            &config,
        )
        .expect("parse arm");
        assert!(parse_out.exists());

        run_command(
            Commands::Summary {
                input: fixture("minimal.docx"),
            },
            &config,
        )
        .expect("summary arm");

        run_command(
            Commands::ProbeFormat {
                input: fixture("minimal.docx"),
                json: true,
                pretty: true,
                output: None,
            },
            &config,
        )
        .expect("probe-format arm");

        run_command(
            Commands::Inventory {
                input: fixture("minimal.docx"),
                json: true,
                pretty: true,
                output: None,
            },
            &config,
        )
        .expect("inventory arm");

        run_command(
            Commands::ListTimes {
                input: list_times_input.clone(),
                json: true,
                pretty: true,
                output: None,
            },
            &config,
        )
        .expect("list-times arm");

        run_command(
            Commands::Manifest {
                input: fixture("minimal.docx"),
                pretty: true,
                output: None,
            },
            &config,
        )
        .expect("manifest arm");

        run_command(
            Commands::InspectMetadata {
                input: metadata_input.clone(),
                json: true,
                pretty: true,
                output: None,
            },
            &config,
        )
        .expect("inspect-metadata arm");

        run_command(
            Commands::InspectDirectory {
                input: inspect_directory_input.clone(),
                json: true,
                pretty: true,
                output: None,
            },
            &config,
        )
        .expect("inspect-directory arm");

        run_command(
            Commands::InspectSectors {
                input: inspect_sectors_input.clone(),
                json: true,
                pretty: true,
                output: None,
            },
            &config,
        )
        .expect("inspect-sectors arm");

        run_command(
            Commands::ReportIndicators {
                input: fixture("minimal.docx"),
                json: true,
                pretty: true,
                output: None,
            },
            &config,
        )
        .expect("report-indicators arm");

        run_command(
            Commands::ExtractLinks {
                input: fixture("minimal.docx"),
                json: true,
                pretty: true,
                output: None,
            },
            &config,
        )
        .expect("extract-links arm");

        run_command(
            Commands::DumpContainer {
                input: fixture("minimal.docx"),
                json: true,
                pretty: true,
                output: None,
            },
            &config,
        )
        .expect("dump-container arm");

        run_command(
            Commands::RecognizeVba {
                input: fixture("minimal.docx"),
                include_source: false,
                json: true,
                pretty: true,
                output: None,
            },
            &config,
        )
        .expect("recognize-vba arm");

        run_command(
            Commands::Security {
                input: fixture("minimal.docx"),
                json: true,
                verbose: false,
            },
            &config,
        )
        .expect("security arm");

        let rules_out = temp_file("rules");
        run_command(
            Commands::Rules {
                input: fixture("minimal.docx"),
                pretty: true,
                output: Some(rules_out.clone()),
                profile: None,
            },
            &config,
        )
        .expect("rules arm");
        assert!(rules_out.exists());

        let diff_out = temp_file("diff");
        run_command(
            Commands::Diff {
                left: fixture("minimal.docx"),
                right: fixture("minimal.docx"),
                pretty: true,
                output: Some(diff_out.clone()),
            },
            &config,
        )
        .expect("diff arm");
        assert!(diff_out.exists());

        run_command(
            Commands::Coverage {
                input: fixture("minimal.docx"),
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

        let artifacts_out = temp_dir("extract_artifacts");
        run_command(
            Commands::ExtractArtifacts {
                input: fixture("minimal.docx"),
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
        let query_out = temp_file("query");
        run_command(
            Commands::Query {
                input: fixture("minimal.docx"),
                node_type: Some("Paragraph".to_string()),
                contains: Some("Hello".to_string()),
                format: Some("docx".to_string()),
                has_external_refs: None,
                has_macros: None,
                pretty: true,
                output: Some(query_out.clone()),
            },
            &config,
        )
        .expect("query arm");
        assert!(query_out.exists());

        let select_out = temp_file("select");
        run_command(
            Commands::Select {
                input: fixture("minimal.docx"),
                node_type: Some("Paragraph".to_string()),
                contains: None,
                format: None,
                has_external_refs: None,
                has_macros: None,
                pretty: true,
                output: Some(select_out.clone()),
            },
            &config,
        )
        .expect("select arm");
        assert!(select_out.exists());

        let grep_out = temp_file("grep");
        run_command(
            Commands::Grep {
                input: fixture("minimal.docx"),
                pattern: "Hello".to_string(),
                node_type: None,
                format: None,
                pretty: true,
                output: Some(grep_out.clone()),
            },
            &config,
        )
        .expect("grep arm");
        assert!(grep_out.exists());

        let extract_out = temp_file("extract");
        run_command(
            Commands::Extract {
                input: fixture("minimal.docx"),
                node_id: Vec::new(),
                node_type: Some("Paragraph".to_string()),
                pretty: true,
                output: Some(extract_out.clone()),
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
                input: fixture("minimal.docx"),
                node_id: "invalid".to_string(),
                format: OutputFormat::Json,
            },
            &config,
        )
        .expect_err("invalid dump-node id should fail");
        assert!(err.to_string().contains("Invalid node ID format"));
    }
}
