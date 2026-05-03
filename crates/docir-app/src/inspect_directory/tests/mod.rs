pub(super) use super::{
    annotate_incoming_references, build_anomaly_counts, build_anomaly_severity_counts,
    build_dead_reference_counts, build_entry_anomaly_tags, build_pointer_counts,
    build_reference_counts, build_tree_density_counts, inspect_directory_bytes, DirectoryEntry,
    EntryAnomalyContext,
};
pub(super) use crate::test_support::{build_test_cfb, build_test_cfb_with_times};

mod annotate_incoming_refs;
mod bucket_counts;

#[test]
fn inspect_directory_reads_cfb_entries() {
    let inspection = inspect_directory_bytes(&build_test_cfb(&[
        ("WordDocument", b"doc"),
        ("VBA/PROJECT", b"meta"),
        ("ObjectPool/1/Ole10Native", b"payload"),
    ]))
    .expect("inspection");

    assert_eq!(inspection.entry_count, 8);
    assert!(inspection
        .role_counts
        .iter()
        .any(|entry| entry.bucket == "state:normal" && entry.count >= 1));
    assert!(inspection
        .role_counts
        .iter()
        .any(|entry| entry.bucket == "classification:word-main-stream" && entry.count == 1));
    assert!(inspection
        .entries
        .iter()
        .any(|entry| entry.path == "WordDocument"
            && entry.classification == "word-main-stream"
            && entry.name_len_raw > 0
            && entry.color_flag_raw == 0
            && entry.object_type_raw == 2));
    assert!(inspection
        .entries
        .iter()
        .any(|entry| entry.path == "Root Entry"
            && entry.child.is_some()
            && entry.child_raw == 1
            && entry.entry_index == 0
            && entry.object_type_raw == 5));
    assert!(inspection
        .entries
        .iter()
        .any(|entry| entry.path == "VBA" && entry.classification == "vba-storage"));
    assert!(inspection
        .entries
        .iter()
        .any(|entry| entry.path == "ObjectPool/1/Ole10Native"
            && entry.classification == "ole-native-payload"));
}

#[test]
fn inspect_directory_preserves_cfb_timestamps() {
    let inspection = inspect_directory_bytes(&build_test_cfb_with_times(
        &[("WordDocument", b"doc")],
        &[("WordDocument", 11, 22)],
    ))
    .expect("inspection");
    let word = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "WordDocument")
        .expect("word");
    assert_eq!(word.created_filetime, Some(11));
    assert_eq!(word.modified_filetime, Some(22));
}

#[test]
fn build_anomaly_counts_detects_structural_anomalies() {
    let anomalies = build_anomaly_counts(&[
        DirectoryEntry {
            entry_index: 1,
            path: "<free>".to_string(),
            entry_type: "unknown".to_string(),
            name_len_raw: 0,
            object_type_raw: 0,
            color_flag_raw: 0,
            state: "free".to_string(),
            classification: "free-slot".to_string(),
            anomaly_severity: "low".to_string(),
            anomaly_tags: vec!["free-slot".to_string(), "unknown-entry-type".to_string()],
            short_cycles: Vec::new(),
            reachable_from_root: false,
            incoming_reference_count: 0,
            incoming_normal_reference_count: 0,
            incoming_anomalous_reference_count: 0,
            incoming_from_root_storage_count: 0,
            incoming_from_storage_count: 0,
            incoming_from_stream_count: 0,
            fanout_count: 0,
            incoming_from: Vec::new(),
            size_bytes: 0,
            start_sector: 0,
            left_sibling_raw: u32::MAX,
            right_sibling_raw: u32::MAX,
            child_raw: u32::MAX,
            left_sibling: None,
            right_sibling: None,
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
        DirectoryEntry {
            entry_index: 2,
            path: "Ghost".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 0,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "orphaned".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "medium".to_string(),
            anomaly_tags: vec!["orphaned-entry".to_string(), "zero-name-length".to_string()],
            short_cycles: Vec::new(),
            reachable_from_root: false,
            incoming_reference_count: 0,
            incoming_normal_reference_count: 0,
            incoming_anomalous_reference_count: 0,
            incoming_from_root_storage_count: 0,
            incoming_from_storage_count: 0,
            incoming_from_stream_count: 0,
            fanout_count: 0,
            incoming_from: Vec::new(),
            size_bytes: 4,
            start_sector: 3,
            left_sibling_raw: u32::MAX,
            right_sibling_raw: u32::MAX,
            child_raw: u32::MAX,
            left_sibling: None,
            right_sibling: None,
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
    ]);

    assert!(anomalies
        .iter()
        .any(|entry| entry.bucket == "free-slot" && entry.count == 1));
    assert!(anomalies
        .iter()
        .any(|entry| entry.bucket == "orphaned-entry" && entry.count == 1));
    assert!(anomalies
        .iter()
        .any(|entry| entry.bucket == "unknown-entry-type" && entry.count == 1));
    assert!(anomalies
        .iter()
        .any(|entry| entry.bucket == "zero-name-length" && entry.count == 1));
}

#[test]
fn build_entry_anomaly_tags_marks_expected_conditions() {
    let tags = build_entry_anomaly_tags(EntryAnomalyContext {
        state: "orphaned",
        entry_type: "unknown",
        name_len_raw: 0,
        entry_index: 2,
        slot_count: 3,
        sector_count: 10,
        start_sector: 0,
        size_bytes: 0,
        left_sibling: None,
        right_sibling: None,
        child: None,
    });
    assert!(tags.iter().any(|tag| tag == "orphaned-entry"));
    assert!(tags.iter().any(|tag| tag == "unknown-entry-type"));
    assert!(tags.iter().any(|tag| tag == "zero-name-length"));

    let free_tags = build_entry_anomaly_tags(EntryAnomalyContext {
        state: "free",
        entry_type: "unknown",
        name_len_raw: 0,
        entry_index: 1,
        slot_count: 3,
        sector_count: 10,
        start_sector: 0,
        size_bytes: 0,
        left_sibling: None,
        right_sibling: None,
        child: None,
    });
    assert!(free_tags.iter().any(|tag| tag == "free-slot"));
    assert!(!free_tags.iter().any(|tag| tag == "zero-name-length"));
}

#[test]
fn build_entry_anomaly_tags_marks_pointer_integrity_issues() {
    let tags = build_entry_anomaly_tags(EntryAnomalyContext {
        state: "normal",
        entry_type: "stream",
        name_len_raw: 12,
        entry_index: 4,
        slot_count: 5,
        sector_count: 10,
        start_sector: 0,
        size_bytes: 8,
        left_sibling: Some(4),
        right_sibling: Some(9),
        child: Some(4),
    });
    assert!(tags.iter().any(|tag| tag == "self-left-sibling"));
    assert!(tags.iter().any(|tag| tag == "self-child"));
    assert!(tags.iter().any(|tag| tag == "dangling-right-sibling"));
}

#[test]
fn build_entry_anomaly_tags_marks_invalid_start_sector() {
    let tags = build_entry_anomaly_tags(EntryAnomalyContext {
        state: "normal",
        entry_type: "stream",
        name_len_raw: 12,
        entry_index: 1,
        slot_count: 8,
        sector_count: 4,
        start_sector: 99,
        size_bytes: 100,
        left_sibling: None,
        right_sibling: None,
        child: None,
    });
    assert!(tags.iter().any(|tag| tag == "invalid-start-sector"));

    let valid_tags = build_entry_anomaly_tags(EntryAnomalyContext {
        state: "normal",
        entry_type: "stream",
        name_len_raw: 12,
        entry_index: 1,
        slot_count: 8,
        sector_count: 10,
        start_sector: 3,
        size_bytes: 100,
        left_sibling: None,
        right_sibling: None,
        child: None,
    });
    assert!(!valid_tags.iter().any(|tag| tag == "invalid-start-sector"));

    let free_tags = build_entry_anomaly_tags(EntryAnomalyContext {
        state: "free",
        entry_type: "unknown",
        name_len_raw: 0,
        entry_index: 2,
        slot_count: 8,
        sector_count: 4,
        start_sector: 99,
        size_bytes: 0,
        left_sibling: None,
        right_sibling: None,
        child: None,
    });
    assert!(!free_tags.iter().any(|tag| tag == "invalid-start-sector"));
}
