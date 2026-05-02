use crate::bucket_count::BucketCount;
use crate::severity::max_severity;

use super::types::{DirectoryAnomalySeverity, DirectoryEntry, EntryAnomalyContext};

pub(super) fn anomaly_severity(tag: &str) -> &'static str {
    match tag {
        "sibling-2-cycle"
        | "child-2-cycle"
        | "sibling-3-cycle"
        | "child-3-cycle"
        | "mixed-2-cycle"
        | "mixed-3-cycle"
        | "incoming-from-anomalous-entry"
        | "dangling-left-sibling"
        | "dangling-right-sibling"
        | "dangling-child"
        | "self-left-sibling"
        | "self-right-sibling"
        | "self-child"
        | "invalid-start-sector" => "high",
        "orphaned-entry"
        | "unknown-entry-type"
        | "multi-referenced-entry"
        | "unreachable-live-entry" => "medium",
        "free-slot" | "zero-name-length" | "unreferenced-live-entry" => "low",
        _ => "medium",
    }
}

pub(super) fn build_entry_anomaly_tags(context: EntryAnomalyContext<'_>) -> Vec<String> {
    let mut tags = Vec::new();
    if context.state == "free" {
        tags.push("free-slot".to_string());
    }
    if context.state == "orphaned" {
        tags.push("orphaned-entry".to_string());
    }
    if context.entry_type == "unknown" {
        tags.push("unknown-entry-type".to_string());
    }
    if context.state != "free" && context.name_len_raw == 0 {
        tags.push("zero-name-length".to_string());
    }
    if context.state != "free"
        && context.entry_type != "unknown"
        && context.size_bytes > 0
        && context.start_sector != u32::MAX
        && context.start_sector != u32::MAX - 1
        && context.start_sector >= context.sector_count
    {
        tags.push("invalid-start-sector".to_string());
    }
    for (label, pointer) in [
        ("dangling-left-sibling", context.left_sibling),
        ("dangling-right-sibling", context.right_sibling),
        ("dangling-child", context.child),
    ] {
        if let Some(pointer) = pointer {
            if pointer >= context.slot_count {
                tags.push(label.to_string());
            }
        }
    }
    if context.left_sibling == Some(context.entry_index) {
        tags.push("self-left-sibling".to_string());
    }
    if context.right_sibling == Some(context.entry_index) {
        tags.push("self-right-sibling".to_string());
    }
    if context.child == Some(context.entry_index) {
        tags.push("self-child".to_string());
    }
    tags
}

pub(super) fn build_anomaly_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        for anomaly in &entry.anomaly_tags {
            *counts.entry(anomaly.clone()).or_insert(0) += 1;
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}

pub(super) fn build_anomaly_catalog(entries: &[DirectoryEntry]) -> Vec<DirectoryAnomalySeverity> {
    let mut seen = std::collections::BTreeSet::<String>::new();
    for entry in entries {
        for anomaly in &entry.anomaly_tags {
            seen.insert(anomaly.clone());
        }
    }
    seen.into_iter()
        .map(|anomaly| DirectoryAnomalySeverity {
            severity: anomaly_severity(&anomaly).to_string(),
            anomaly,
        })
        .collect()
}

pub(super) fn build_anomaly_severity_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        for anomaly in &entry.anomaly_tags {
            *counts
                .entry(anomaly_severity(anomaly).to_string())
                .or_insert(0) += 1;
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}

pub(super) fn build_directory_score(entries: &[DirectoryEntry]) -> String {
    max_severity(
        entries
            .iter()
            .filter(|entry| entry.state != "free")
            .map(|entry| entry.anomaly_severity.clone()),
    )
}

pub(super) fn build_role_counts(entries: &[DirectoryEntry]) -> Vec<BucketCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        *counts.entry(format!("state:{}", entry.state)).or_insert(0) += 1;
        *counts
            .entry(format!("entry-type:{}", entry.entry_type))
            .or_insert(0) += 1;
        *counts
            .entry(format!("classification:{}", entry.classification))
            .or_insert(0) += 1;
    }
    counts
        .into_iter()
        .map(|(bucket, count)| BucketCount { bucket, count })
        .collect()
}
