use crate::bucket_count::BucketCount;

use super::types::DirectoryEntry;

pub(super) fn build_reference_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        let bucket = match entry.incoming_reference_count {
            0 => "incoming:0",
            1 => "incoming:1",
            _ => "incoming:many",
        };
        *counts.entry(bucket.to_string()).or_insert(0) += 1;

        if entry.state == "normal" && entry.entry_type != "root-storage" {
            let live_bucket = match entry.incoming_reference_count {
                0 => "live-incoming:0",
                1 => "live-incoming:1",
                _ => "live-incoming:many",
            };
            *counts.entry(live_bucket.to_string()).or_insert(0) += 1;
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}

pub(super) fn build_pointer_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        for (kind, pointer, dangling_tag) in [
            ("left", entry.left_sibling, "dangling-left-sibling"),
            ("right", entry.right_sibling, "dangling-right-sibling"),
            ("child", entry.child, "dangling-child"),
        ] {
            if pointer.is_some() {
                *counts.entry(format!("{kind}:present")).or_insert(0) += 1;
            }
            if entry.anomaly_tags.iter().any(|tag| tag == dangling_tag) {
                *counts.entry(format!("{kind}:dangling")).or_insert(0) += 1;
            }
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}

pub(super) fn build_tree_density_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        for (kind, pointer) in [
            ("left", entry.left_sibling),
            ("right", entry.right_sibling),
            ("child", entry.child),
        ] {
            if pointer.is_some() {
                *counts
                    .entry(format!("{kind}:state:{}", entry.state))
                    .or_insert(0) += 1;
                *counts
                    .entry(format!("{kind}:entry-type:{}", entry.entry_type))
                    .or_insert(0) += 1;
            }
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}

pub(super) fn build_dangling_state_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        for (tag, kind) in [
            ("dangling-left-sibling", "left"),
            ("dangling-right-sibling", "right"),
            ("dangling-child", "child"),
        ] {
            if entry.anomaly_tags.iter().any(|value| value == tag) {
                *counts
                    .entry(format!("{kind}:state:{}", entry.state))
                    .or_insert(0) += 1;
            }
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}

pub(super) fn build_self_reference_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        for bucket in ["self-left-sibling", "self-right-sibling", "self-child"] {
            if entry.anomaly_tags.iter().any(|value| value == bucket) {
                *counts.entry(bucket.to_string()).or_insert(0) += 1;
            }
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}

pub(super) fn build_short_cycle_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        for bucket in &entry.short_cycles {
            *counts.entry(bucket.clone()).or_insert(0) += 1;
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}

pub(super) fn build_reachability_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        let bucket = if entry.state == "normal" && entry.entry_type != "root-storage" {
            if entry.reachable_from_root {
                "live-reachable"
            } else {
                "live-unreachable"
            }
        } else if entry.reachable_from_root {
            "non-live-reachable"
        } else {
            "non-live-unreachable"
        };
        *counts.entry(bucket.to_string()).or_insert(0) += 1;
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}

pub(super) fn build_incoming_source_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        *counts
            .entry("incoming:state:normal".to_string())
            .or_insert(0) += entry.incoming_normal_reference_count;
        *counts
            .entry("incoming:state:anomalous".to_string())
            .or_insert(0) += entry.incoming_anomalous_reference_count;
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}

pub(super) fn build_incoming_source_type_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        *counts
            .entry("incoming:source-type:root-storage".to_string())
            .or_insert(0) += entry.incoming_from_root_storage_count;
        *counts
            .entry("incoming:source-type:storage".to_string())
            .or_insert(0) += entry.incoming_from_storage_count;
        *counts
            .entry("incoming:source-type:stream".to_string())
            .or_insert(0) += entry.incoming_from_stream_count;
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}

pub(super) fn build_dead_reference_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        if (entry.state == "free" || entry.state == "orphaned")
            && entry.incoming_reference_count > 0
        {
            *counts
                .entry(format!("dead-reference:state:{}", entry.state))
                .or_insert(0) += 1;
            *counts
                .entry("dead-reference:source-type:root-storage".to_string())
                .or_insert(0) += entry.incoming_from_root_storage_count;
            *counts
                .entry("dead-reference:source-type:storage".to_string())
                .or_insert(0) += entry.incoming_from_storage_count;
            *counts
                .entry("dead-reference:source-type:stream".to_string())
                .or_insert(0) += entry.incoming_from_stream_count;
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}

pub(super) fn build_fanout_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        *counts
            .entry(format!("fanout:{}", entry.fanout_count))
            .or_insert(0) += 1;
        if entry.state == "normal" && entry.entry_type != "root-storage" {
            *counts
                .entry(format!("live-fanout:{}", entry.fanout_count))
                .or_insert(0) += 1;
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}