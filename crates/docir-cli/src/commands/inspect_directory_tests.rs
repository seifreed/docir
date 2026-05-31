use super::{format_inspection_text, run};
use crate::cli::JsonOutputOpts;
use crate::test_support;
use docir_app::{
    BucketCount, DirectoryAnomalySeverity, DirectoryEntry, DirectoryInspection, ParserConfig,
    test_support::{build_test_cfb, build_test_cfb_with_times},
};
use std::fs;

#[test]
fn inspect_directory_run_writes_json() {
    let input = test_support::temp_file("legacy", "doc");
    let output = test_support::temp_file("legacy", "json");
    fs::write(
        &input,
        build_test_cfb(&[("WordDocument", b"doc"), ("VBA/PROJECT", b"meta")]),
    )
    .expect("fixture");

    run(
        input.clone(),
        JsonOutputOpts {
            json: true,
            pretty: true,
            output: Some(output.clone()),
        },
        &ParserConfig::default(),
    )
    .expect("inspect-directory json");

    let text = fs::read_to_string(&output).expect("output");
    assert!(text.contains("\"container\": \"cfb-ole\""));
    assert!(text.contains("\"role_counts\": ["));
    assert!(text.contains("\"path\": \"WordDocument\""));
    assert!(text.contains("\"entry_index\": 1"));
    assert!(text.contains("\"name_len_raw\": 26"));
    assert!(text.contains("\"object_type_raw\": 2"));
    assert!(text.contains("\"color_flag_raw\": 0"));
    assert!(text.contains("\"left_sibling_raw\": 4294967295"));
    assert!(text.contains("\"child_raw\": 4294967295"));
    assert!(text.contains("\"classification\": \"word-main-stream\""));
    assert!(text.contains("\"anomaly_counts\": ["));
    assert!(text.contains("\"child\":"));

    let _ = fs::remove_file(input);
    let _ = fs::remove_file(output);
}

#[test]
fn inspect_directory_run_writes_text() {
    let input = test_support::temp_file("legacy_text", "doc");
    let output = test_support::temp_file("legacy_text", "txt");
    fs::write(
        &input,
        build_test_cfb_with_times(&[("WordDocument", b"doc")], &[("WordDocument", 11, 22)]),
    )
    .expect("fixture");

    run(
        input.clone(),
        JsonOutputOpts {
            json: false,
            pretty: false,
            output: Some(output.clone()),
        },
        &ParserConfig::default(),
    )
    .expect("inspect-directory text");

    let text = fs::read_to_string(&output).expect("output");
    assert!(text.contains("Role Counts:"));
    assert!(text.contains("classification:word-main-stream: 1"));
    assert!(text.contains("Anomalies:"));
    assert!(text.contains("free-slot:"));
    assert!(text.contains("Reference Summary:"));
    assert!(text.contains("incoming:"));
    assert!(text.contains("Pointer Summary:"));
    assert!(text.contains("Tree Density Summary:"));
    assert!(text.contains("Reachability Summary:"));
    assert!(text.contains("Incoming Source Summary:"));
    assert!(text.contains("Directory:"));
    assert!(text.contains("stream: WordDocument"));
    assert!(text.contains("Entry Index: 1"));
    assert!(text.contains("Name Length Raw: 26"));
    assert!(text.contains("Object Type Raw: 2"));
    assert!(text.contains("Color Flag Raw: 0"));
    assert!(text.contains("Left Sibling Raw: 4294967295"));
    assert!(text.contains("Child Raw: 4294967295"));
    assert!(text.contains("Classification: word-main-stream"));
    assert!(text.contains("Anomaly Tags:"));
    assert!(
        text.contains("zero-name-length")
            || text.contains("orphaned-entry")
            || text.contains("free-slot")
    );
    assert!(text.contains("Child:"));
    assert!(text.contains("Created: 11"));

    let _ = fs::remove_file(input);
    let _ = fs::remove_file(output);
}

#[test]
fn format_inspection_text_renders_expected_fields() {
    let text = format_inspection_text(&directory_inspection_fixture());

    assert_inspection_summary_text(&text);
    assert_inspection_bucket_text(&text);
    assert_directory_entry_text(&text);
}

fn directory_inspection_fixture() -> DirectoryInspection {
    DirectoryInspection {
        container: "cfb-ole".to_string(),
        entry_count: 1,
        directory_score: "medium".to_string(),
        role_counts: vec![
            bucket_count("state:normal", 1),
            bucket_count("classification:word-main-stream", 1),
        ],
        anomaly_counts: vec![bucket_count("orphaned-entry", 1)],
        anomaly_catalog: vec![DirectoryAnomalySeverity {
            anomaly: "orphaned-entry".to_string(),
            severity: "medium".to_string(),
        }],
        anomaly_severity_counts: vec![bucket_count("medium", 1)],
        reference_counts: vec![
            bucket_count("incoming:many", 1),
            bucket_count("live-incoming:many", 1),
        ],
        pointer_counts: vec![
            bucket_count("right:present", 1),
            bucket_count("right:dangling", 1),
        ],
        tree_density_counts: vec![
            bucket_count("right:state:normal", 1),
            bucket_count("right:entry-type:stream", 1),
        ],
        dangling_state_counts: vec![bucket_count("right:state:normal", 1)],
        self_reference_counts: vec![bucket_count("self-right-sibling", 1)],
        short_cycle_counts: vec![bucket_count("sibling-2-cycle", 1)],
        reachability_counts: vec![bucket_count("live-reachable", 1)],
        incoming_source_counts: vec![bucket_count("incoming:state:normal", 2)],
        incoming_source_type_counts: vec![
            bucket_count("incoming:source-type:root-storage", 1),
            bucket_count("incoming:source-type:storage", 1),
        ],
        dead_reference_counts: vec![bucket_count("dead-reference:state:orphaned", 1)],
        fanout_counts: vec![bucket_count("fanout:0", 1), bucket_count("fanout:2", 1)],
        entries: vec![directory_entry_fixture()],
    }
}

fn directory_entry_fixture() -> DirectoryEntry {
    DirectoryEntry {
        entry_index: 1,
        path: "WordDocument".to_string(),
        entry_type: "stream".to_string(),
        name_len_raw: 26,
        object_type_raw: 2,
        color_flag_raw: 0,
        state: "normal".to_string(),
        classification: "word-main-stream".to_string(),
        anomaly_severity: "medium".to_string(),
        anomaly_tags: vec!["orphaned-entry".to_string()],
        short_cycles: vec!["sibling-2-cycle".to_string()],
        reachable_from_root: true,
        fanout_count: 2,
        incoming_reference_count: 2,
        incoming_normal_reference_count: 2,
        incoming_anomalous_reference_count: 0,
        incoming_from_root_storage_count: 1,
        incoming_from_storage_count: 1,
        incoming_from_stream_count: 0,
        incoming_from: vec!["left:Root Entry#0".to_string(), "right:VBA#2".to_string()],
        size_bytes: 3,
        start_sector: 0,
        left_sibling_raw: u32::MAX,
        right_sibling_raw: 2,
        child_raw: u32::MAX,
        left_sibling: None,
        right_sibling: Some(2),
        child: None,
        created_filetime: None,
        modified_filetime: None,
    }
}

fn bucket_count(bucket: &str, count: usize) -> BucketCount {
    BucketCount {
        bucket: bucket.to_string(),
        count,
    }
}

fn assert_inspection_summary_text(text: &str) {
    assert_contains_all(
        text,
        &[
            "Entries: 1",
            "Directory Score: medium",
            "Role Counts:",
            "Anomalies:",
            "Anomaly Severity Catalog:",
            "Anomaly Severity Summary:",
            "Reference Summary:",
            "Pointer Summary:",
            "Tree Density Summary:",
            "Dangling By State:",
            "Self References:",
            "Short Cycles:",
            "Reachability Summary:",
            "Incoming Source Summary:",
            "Incoming Source Types:",
            "Dead But Referenced:",
            "Fanout Summary:",
        ],
    );
}

fn assert_inspection_bucket_text(text: &str) {
    assert_contains_all(
        text,
        &[
            "classification:word-main-stream: 1",
            "orphaned-entry: 1",
            "orphaned-entry: medium",
            "medium: 1",
            "incoming:many: 1",
            "right:present: 1",
            "right:state:normal: 1",
        ],
    );
}

fn assert_directory_entry_text(text: &str) {
    assert_contains_all(
        text,
        &[
            "Entry Index: 1",
            "Name Length Raw: 26",
            "Object Type Raw: 2",
            "Color Flag Raw: 0",
            "State: normal",
            "Anomaly Severity: medium",
            "Short Cycles: sibling-2-cycle",
            "Reachable From Root: true",
            "Fanout: 2",
            "Incoming References: 2",
            "Incoming From Normal: 2",
            "Incoming From Anomalous: 0",
            "Incoming From Root Storage: 1",
            "Incoming From Storage: 1",
            "Incoming From Stream: 0",
            "Incoming From: left:Root Entry#0, right:VBA#2",
            "Sector: 0",
            "Anomaly Tags: orphaned-entry",
            "Left Sibling Raw: 4294967295",
            "Right Sibling Raw: 2",
            "Right Sibling: 2",
        ],
    );
}

fn assert_contains_all(text: &str, expected: &[&str]) {
    for fragment in expected {
        assert!(text.contains(fragment), "missing fragment: {fragment}");
    }
}
