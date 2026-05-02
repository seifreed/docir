use super::inspect_directory_bytes;
use super::{
    annotate_incoming_references, build_anomaly_counts, build_anomaly_severity_counts,
    build_dead_reference_counts, build_entry_anomaly_tags, build_pointer_counts,
    build_reference_counts, build_tree_density_counts, DirectoryEntry, EntryAnomalyContext,
};
use crate::test_support::{build_test_cfb, build_test_cfb_with_times};

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

#[test]
fn build_reference_counts_groups_incoming_buckets() {
    let counts = build_reference_counts(&[
        DirectoryEntry {
            entry_index: 0,
            path: "Root Entry".to_string(),
            entry_type: "root-storage".to_string(),
            name_len_raw: 20,
            object_type_raw: 5,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "root-storage".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
            short_cycles: Vec::new(),
            reachable_from_root: true,
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
            child_raw: 1,
            left_sibling: None,
            right_sibling: None,
            child: Some(1),
            created_filetime: None,
            modified_filetime: None,
        },
        DirectoryEntry {
            entry_index: 1,
            path: "A".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
            short_cycles: Vec::new(),
            reachable_from_root: true,
            incoming_reference_count: 1,
            incoming_normal_reference_count: 1,
            incoming_anomalous_reference_count: 0,
            incoming_from_root_storage_count: 1,
            incoming_from_storage_count: 0,
            incoming_from_stream_count: 0,
            fanout_count: 0,
            incoming_from: vec!["child:Root Entry#0".to_string()],
            size_bytes: 1,
            start_sector: 1,
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
            path: "B".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "medium".to_string(),
            anomaly_tags: vec!["multi-referenced-entry".to_string()],
            short_cycles: Vec::new(),
            reachable_from_root: true,
            incoming_reference_count: 3,
            incoming_normal_reference_count: 3,
            incoming_anomalous_reference_count: 0,
            incoming_from_root_storage_count: 0,
            incoming_from_storage_count: 0,
            incoming_from_stream_count: 3,
            fanout_count: 0,
            incoming_from: vec![
                "left:A#1".to_string(),
                "right:C#3".to_string(),
                "child:D#4".to_string(),
            ],
            size_bytes: 1,
            start_sector: 2,
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
    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "incoming:0" && entry.count == 1));
    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "incoming:1" && entry.count == 1));
    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "incoming:many" && entry.count == 1));
    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "live-incoming:1" && entry.count == 1));
    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "live-incoming:many" && entry.count == 1));
}

#[test]
fn build_pointer_counts_groups_present_and_dangling_pointers() {
    let counts = build_pointer_counts(&[
        DirectoryEntry {
            entry_index: 1,
            path: "A".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "high".to_string(),
            anomaly_tags: vec!["dangling-right-sibling".to_string()],
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
            size_bytes: 1,
            start_sector: 1,
            left_sibling_raw: u32::MAX,
            right_sibling_raw: 9,
            child_raw: 2,
            left_sibling: None,
            right_sibling: Some(9),
            child: Some(2),
            created_filetime: None,
            modified_filetime: None,
        },
        DirectoryEntry {
            entry_index: 2,
            path: "B".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "high".to_string(),
            anomaly_tags: vec!["dangling-child".to_string()],
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
            size_bytes: 1,
            start_sector: 2,
            left_sibling_raw: 1,
            right_sibling_raw: u32::MAX,
            child_raw: 8,
            left_sibling: Some(1),
            right_sibling: None,
            child: Some(8),
            created_filetime: None,
            modified_filetime: None,
        },
    ]);

    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "left:present" && entry.count == 1));
    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "right:present" && entry.count == 1));
    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "child:present" && entry.count == 2));
    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "right:dangling" && entry.count == 1));
    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "child:dangling" && entry.count == 1));
}

#[test]
fn build_tree_density_counts_groups_pointer_usage_by_state_and_entry_type() {
    let counts = build_tree_density_counts(&[
        DirectoryEntry {
            entry_index: 1,
            path: "A".to_string(),
            entry_type: "storage".to_string(),
            name_len_raw: 4,
            object_type_raw: 1,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "storage".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
            short_cycles: Vec::new(),
            reachable_from_root: true,
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
            right_sibling_raw: 2,
            child_raw: 3,
            left_sibling: None,
            right_sibling: Some(2),
            child: Some(3),
            created_filetime: None,
            modified_filetime: None,
        },
        DirectoryEntry {
            entry_index: 2,
            path: "B".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "orphaned".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
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
            size_bytes: 1,
            start_sector: 2,
            left_sibling_raw: 1,
            right_sibling_raw: u32::MAX,
            child_raw: u32::MAX,
            left_sibling: Some(1),
            right_sibling: None,
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
    ]);

    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "right:state:normal" && entry.count == 1));
    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "child:entry-type:storage" && entry.count == 1));
    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "left:state:orphaned" && entry.count == 1));
    assert!(counts
        .iter()
        .any(|entry| entry.bucket == "left:entry-type:stream" && entry.count == 1));
}

#[test]
fn annotate_incoming_references_marks_unreferenced_and_multi_referenced_entries() {
    let mut entries = vec![
        DirectoryEntry {
            entry_index: 0,
            path: "Root Entry".to_string(),
            entry_type: "root-storage".to_string(),
            name_len_raw: 20,
            object_type_raw: 5,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "root-storage".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
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
            child_raw: 1,
            left_sibling: None,
            right_sibling: None,
            child: Some(1),
            created_filetime: None,
            modified_filetime: None,
        },
        DirectoryEntry {
            entry_index: 1,
            path: "A".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
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
            size_bytes: 1,
            start_sector: 1,
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
            path: "B".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
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
            size_bytes: 1,
            start_sector: 2,
            left_sibling_raw: 1,
            right_sibling_raw: 1,
            child_raw: u32::MAX,
            left_sibling: Some(1),
            right_sibling: Some(1),
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
    ];

    annotate_incoming_references(&mut entries);

    let a = entries.iter().find(|entry| entry.entry_index == 1).unwrap();
    let b = entries.iter().find(|entry| entry.entry_index == 2).unwrap();
    assert_eq!(a.incoming_reference_count, 3);
    assert!(a
        .anomaly_tags
        .iter()
        .any(|tag| tag == "multi-referenced-entry"));
    assert!(b
        .anomaly_tags
        .iter()
        .any(|tag| tag == "unreferenced-live-entry"));
}

#[test]
fn annotate_incoming_references_marks_unreachable_and_incoming_source_quality() {
    let mut entries = vec![
        DirectoryEntry {
            entry_index: 0,
            path: "Root Entry".to_string(),
            entry_type: "root-storage".to_string(),
            name_len_raw: 20,
            object_type_raw: 5,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "root-storage".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
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
            entry_index: 1,
            path: "Live".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 10,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
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
            start_sector: 1,
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
            name_len_raw: 10,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "orphaned".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
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
            start_sector: 2,
            left_sibling_raw: u32::MAX,
            right_sibling_raw: 1,
            child_raw: u32::MAX,
            left_sibling: None,
            right_sibling: Some(1),
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
    ];

    annotate_incoming_references(&mut entries);
    let live = entries.iter().find(|entry| entry.entry_index == 1).unwrap();
    assert!(!live.reachable_from_root);
    assert_eq!(live.incoming_reference_count, 1);
    assert_eq!(live.incoming_normal_reference_count, 0);
    assert_eq!(live.incoming_anomalous_reference_count, 1);
    assert!(live
        .anomaly_tags
        .iter()
        .any(|tag| tag == "unreachable-live-entry"));
    assert!(live
        .anomaly_tags
        .iter()
        .any(|tag| tag == "incoming-from-anomalous-entry"));
}

#[test]
fn annotate_incoming_references_marks_short_cycles() {
    let mut entries = vec![
        DirectoryEntry {
            entry_index: 0,
            path: "Root Entry".to_string(),
            entry_type: "root-storage".to_string(),
            name_len_raw: 20,
            object_type_raw: 5,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "root-storage".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
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
            child_raw: 1,
            left_sibling: None,
            right_sibling: None,
            child: Some(1),
            created_filetime: None,
            modified_filetime: None,
        },
        DirectoryEntry {
            entry_index: 1,
            path: "A".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
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
            size_bytes: 1,
            start_sector: 1,
            left_sibling_raw: 2,
            right_sibling_raw: u32::MAX,
            child_raw: u32::MAX,
            left_sibling: Some(2),
            right_sibling: None,
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
        DirectoryEntry {
            entry_index: 2,
            path: "B".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
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
            size_bytes: 1,
            start_sector: 2,
            left_sibling_raw: u32::MAX,
            right_sibling_raw: 1,
            child_raw: u32::MAX,
            left_sibling: None,
            right_sibling: Some(1),
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
    ];

    annotate_incoming_references(&mut entries);
    let a = entries.iter().find(|entry| entry.entry_index == 1).unwrap();
    assert!(a.short_cycles.iter().any(|tag| tag == "sibling-2-cycle"));
    assert!(a.anomaly_tags.iter().any(|tag| tag == "sibling-2-cycle"));
}

#[test]
fn annotate_incoming_references_marks_mixed_cycles_and_dead_sources() {
    let mut entries = vec![
        DirectoryEntry {
            entry_index: 0,
            path: "Root Entry".to_string(),
            entry_type: "root-storage".to_string(),
            name_len_raw: 20,
            object_type_raw: 5,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "root-storage".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
            short_cycles: Vec::new(),
            reachable_from_root: false,
            fanout_count: 0,
            incoming_reference_count: 0,
            incoming_normal_reference_count: 0,
            incoming_anomalous_reference_count: 0,
            incoming_from_root_storage_count: 0,
            incoming_from_storage_count: 0,
            incoming_from_stream_count: 0,
            incoming_from: Vec::new(),
            size_bytes: 0,
            start_sector: 0,
            left_sibling_raw: u32::MAX,
            right_sibling_raw: u32::MAX,
            child_raw: 1,
            left_sibling: None,
            right_sibling: None,
            child: Some(1),
            created_filetime: None,
            modified_filetime: None,
        },
        DirectoryEntry {
            entry_index: 1,
            path: "A".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
            short_cycles: Vec::new(),
            reachable_from_root: false,
            fanout_count: 0,
            incoming_reference_count: 0,
            incoming_normal_reference_count: 0,
            incoming_anomalous_reference_count: 0,
            incoming_from_root_storage_count: 0,
            incoming_from_storage_count: 0,
            incoming_from_stream_count: 0,
            incoming_from: Vec::new(),
            size_bytes: 1,
            start_sector: 1,
            left_sibling_raw: u32::MAX,
            right_sibling_raw: u32::MAX,
            child_raw: 2,
            left_sibling: None,
            right_sibling: None,
            child: Some(2),
            created_filetime: None,
            modified_filetime: None,
        },
        DirectoryEntry {
            entry_index: 2,
            path: "Ghost".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "orphaned".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
            short_cycles: Vec::new(),
            reachable_from_root: false,
            fanout_count: 0,
            incoming_reference_count: 0,
            incoming_normal_reference_count: 0,
            incoming_anomalous_reference_count: 0,
            incoming_from_root_storage_count: 0,
            incoming_from_storage_count: 0,
            incoming_from_stream_count: 0,
            incoming_from: Vec::new(),
            size_bytes: 1,
            start_sector: 2,
            left_sibling_raw: u32::MAX,
            right_sibling_raw: 1,
            child_raw: u32::MAX,
            left_sibling: None,
            right_sibling: Some(1),
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
    ];

    annotate_incoming_references(&mut entries);
    let live = entries.iter().find(|entry| entry.entry_index == 1).unwrap();
    assert!(live.short_cycles.iter().any(|tag| tag == "mixed-2-cycle"));
    let dead_counts = build_dead_reference_counts(&entries);
    assert!(dead_counts
        .iter()
        .any(|entry| entry.bucket == "dead-reference:state:orphaned" && entry.count == 1));
    assert!(dead_counts
        .iter()
        .any(|entry| entry.bucket == "dead-reference:source-type:stream" && entry.count == 1));
    let severity_counts = build_anomaly_severity_counts(&entries);
    assert!(severity_counts
        .iter()
        .any(|entry| entry.bucket == "high" && entry.count >= 1));
}

#[test]
fn annotate_incoming_references_marks_mixed_three_cycles() {
    let mut entries = vec![
        DirectoryEntry {
            entry_index: 0,
            path: "Root Entry".to_string(),
            entry_type: "root-storage".to_string(),
            name_len_raw: 20,
            object_type_raw: 5,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "root-storage".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
            short_cycles: Vec::new(),
            reachable_from_root: false,
            fanout_count: 0,
            incoming_reference_count: 0,
            incoming_normal_reference_count: 0,
            incoming_anomalous_reference_count: 0,
            incoming_from_root_storage_count: 0,
            incoming_from_storage_count: 0,
            incoming_from_stream_count: 0,
            incoming_from: Vec::new(),
            size_bytes: 0,
            start_sector: 0,
            left_sibling_raw: u32::MAX,
            right_sibling_raw: u32::MAX,
            child_raw: 1,
            left_sibling: None,
            right_sibling: None,
            child: Some(1),
            created_filetime: None,
            modified_filetime: None,
        },
        DirectoryEntry {
            entry_index: 1,
            path: "A".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
            short_cycles: Vec::new(),
            reachable_from_root: false,
            fanout_count: 0,
            incoming_reference_count: 0,
            incoming_normal_reference_count: 0,
            incoming_anomalous_reference_count: 0,
            incoming_from_root_storage_count: 0,
            incoming_from_storage_count: 0,
            incoming_from_stream_count: 0,
            incoming_from: Vec::new(),
            size_bytes: 1,
            start_sector: 1,
            left_sibling_raw: u32::MAX,
            right_sibling_raw: 2,
            child_raw: u32::MAX,
            left_sibling: None,
            right_sibling: Some(2),
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
        DirectoryEntry {
            entry_index: 2,
            path: "B".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
            short_cycles: Vec::new(),
            reachable_from_root: false,
            fanout_count: 0,
            incoming_reference_count: 0,
            incoming_normal_reference_count: 0,
            incoming_anomalous_reference_count: 0,
            incoming_from_root_storage_count: 0,
            incoming_from_storage_count: 0,
            incoming_from_stream_count: 0,
            incoming_from: Vec::new(),
            size_bytes: 1,
            start_sector: 2,
            left_sibling_raw: u32::MAX,
            right_sibling_raw: u32::MAX,
            child_raw: 3,
            left_sibling: None,
            right_sibling: None,
            child: Some(3),
            created_filetime: None,
            modified_filetime: None,
        },
        DirectoryEntry {
            entry_index: 3,
            path: "C".to_string(),
            entry_type: "stream".to_string(),
            name_len_raw: 4,
            object_type_raw: 2,
            color_flag_raw: 0,
            state: "normal".to_string(),
            classification: "stream".to_string(),
            anomaly_severity: "none".to_string(),
            anomaly_tags: Vec::new(),
            short_cycles: Vec::new(),
            reachable_from_root: false,
            fanout_count: 0,
            incoming_reference_count: 0,
            incoming_normal_reference_count: 0,
            incoming_anomalous_reference_count: 0,
            incoming_from_root_storage_count: 0,
            incoming_from_storage_count: 0,
            incoming_from_stream_count: 0,
            incoming_from: Vec::new(),
            size_bytes: 1,
            start_sector: 3,
            left_sibling_raw: 1,
            right_sibling_raw: u32::MAX,
            child_raw: u32::MAX,
            left_sibling: Some(1),
            right_sibling: None,
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
    ];

    annotate_incoming_references(&mut entries);
    let a = entries.iter().find(|entry| entry.entry_index == 1).unwrap();
    assert!(a.short_cycles.iter().any(|tag| tag == "mixed-3-cycle"));
    assert!(a.anomaly_tags.iter().any(|tag| tag == "mixed-3-cycle"));
    assert_eq!(a.anomaly_severity, "high");
}
