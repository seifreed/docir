use super::{
    build_pointer_counts, build_reference_counts, build_tree_density_counts, DirectoryEntry,
};

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
