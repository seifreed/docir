use super::dispatch;
use crate::test_support;
use crate::{
    Commands, CoverageExportFormat, CoverageExportMode, JsonOutputOpts, OutputFormat,
    PrettyOutputOpts,
};
use docir_app::{
    ParserConfig,
    test_support::{
        TestPropertyValue, build_test_cfb, build_test_cfb_with_times,
        build_test_property_set_stream,
    },
};
use std::fs;
use std::path::PathBuf;

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

fn json_stdout_opts() -> JsonOutputOpts {
    JsonOutputOpts {
        json: true,
        pretty: true,
        output: None,
    }
}

fn list_times_fixture() -> PathBuf {
    let list_times_input = test_support::temp_file("times", "doc");
    fs::write(
        &list_times_input,
        build_test_cfb_with_times(&[("WordDocument", b"doc")], &[("WordDocument", 10, 20)]),
    )
    .expect("list-times fixture");
    list_times_input
}

fn metadata_fixture() -> PathBuf {
    let metadata_input = test_support::temp_file("metadata", "doc");
    fs::write(
        &metadata_input,
        build_test_cfb(&[(
            "\u{0005}SummaryInformation",
            &build_test_property_set_stream(&[(2, TestPropertyValue::Str("Specimen"))]),
        )]),
    )
    .expect("metadata fixture");
    metadata_input
}

fn inspect_directory_fixture() -> PathBuf {
    let inspect_directory_input = test_support::temp_file("inspect_directory", "doc");
    fs::write(
        &inspect_directory_input,
        build_test_cfb(&[
            ("WordDocument", b"doc"),
            ("VBA/PROJECT", b"meta"),
            ("ObjectPool/1/Ole10Native", b"payload"),
        ]),
    )
    .expect("inspect-directory fixture");
    inspect_directory_input
}

fn inspect_sectors_fixture() -> PathBuf {
    let inspect_sectors_input = test_support::temp_file("inspect_sectors", "doc");
    fs::write(
        &inspect_sectors_input,
        build_test_cfb(&[("WordDocument", b"doc"), ("VBA/PROJECT", b"meta")]),
    )
    .expect("inspect-sectors fixture");
    inspect_sectors_input
}

#[test]
fn dispatch_routes_inspect_commands() {
    let config = ParserConfig::default();
    let list_times_input = list_times_fixture();
    let metadata_input = metadata_fixture();
    let inspect_directory_input = inspect_directory_fixture();
    let inspect_sectors_input = inspect_sectors_fixture();
    dispatch(
        Commands::ListTimes {
            input: list_times_input.clone(),
            output_opts: json_stdout_opts(),
        },
        &config,
    )
    .expect("list-times arm");

    dispatch(
        Commands::InspectMetadata {
            input: metadata_input.clone(),
            output_opts: json_stdout_opts(),
        },
        &config,
    )
    .expect("inspect-metadata arm");

    dispatch(
        Commands::InspectDirectory {
            input: inspect_directory_input.clone(),
            output_opts: json_stdout_opts(),
        },
        &config,
    )
    .expect("inspect-directory arm");

    dispatch(
        Commands::InspectSectors {
            input: inspect_sectors_input.clone(),
            output_opts: json_stdout_opts(),
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
