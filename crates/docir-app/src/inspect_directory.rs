use crate::{AppResult, ParserConfig};
use docir_parser::ole::{Cfb, CfbDirectoryState, CfbEntryType};
use docir_parser::ParseError as ParserParseError;
use serde::Serialize;
use std::fs;
use std::path::Path;

/// Dedicated structural view of normal CFB directory entries.
#[derive(Debug, Clone, Serialize)]
pub struct DirectoryInspection {
    pub container: String,
    pub entry_count: usize,
    pub directory_score: String,
    pub role_counts: Vec<DirectoryRoleCount>,
    pub anomaly_counts: Vec<DirectoryAnomalyCount>,
    pub anomaly_catalog: Vec<DirectoryAnomalySeverity>,
    pub anomaly_severity_counts: Vec<DirectoryAnomalySeverityCount>,
    pub reference_counts: Vec<DirectoryReferenceCount>,
    pub pointer_counts: Vec<DirectoryPointerCount>,
    pub tree_density_counts: Vec<DirectoryTreeDensityCount>,
    pub dangling_state_counts: Vec<DirectoryPointerStateCount>,
    pub self_reference_counts: Vec<DirectorySelfReferenceCount>,
    pub short_cycle_counts: Vec<DirectoryCycleCount>,
    pub reachability_counts: Vec<DirectoryReachabilityCount>,
    pub incoming_source_counts: Vec<DirectoryIncomingSourceCount>,
    pub incoming_source_type_counts: Vec<DirectoryIncomingSourceTypeCount>,
    pub dead_reference_counts: Vec<DirectoryDeadReferenceCount>,
    pub fanout_counts: Vec<DirectoryFanoutCount>,
    pub entries: Vec<DirectoryEntry>,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryRoleCount {
    pub role: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryAnomalyCount {
    pub anomaly: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryAnomalySeverity {
    pub anomaly: String,
    pub severity: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryAnomalySeverityCount {
    pub severity: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryReferenceCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryPointerCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryTreeDensityCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryPointerStateCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectorySelfReferenceCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryCycleCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryReachabilityCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryIncomingSourceCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryIncomingSourceTypeCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryDeadReferenceCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DirectoryFanoutCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Copy)]
struct EntryAnomalyContext<'a> {
    state: &'a str,
    entry_type: &'a str,
    name_len_raw: u16,
    entry_index: u32,
    slot_count: u32,
    sector_count: u32,
    start_sector: u32,
    size_bytes: u64,
    left_sibling: Option<u32>,
    right_sibling: Option<u32>,
    child: Option<u32>,
}

/// One CFB directory entry surfaced for analyst inspection.
#[derive(Debug, Clone, Serialize)]
pub struct DirectoryEntry {
    pub entry_index: u32,
    pub path: String,
    pub entry_type: String,
    pub name_len_raw: u16,
    pub object_type_raw: u8,
    pub color_flag_raw: u8,
    pub state: String,
    pub classification: String,
    pub anomaly_severity: String,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub anomaly_tags: Vec<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub short_cycles: Vec<String>,
    pub reachable_from_root: bool,
    pub fanout_count: usize,
    pub incoming_reference_count: usize,
    pub incoming_normal_reference_count: usize,
    pub incoming_anomalous_reference_count: usize,
    pub incoming_from_root_storage_count: usize,
    pub incoming_from_storage_count: usize,
    pub incoming_from_stream_count: usize,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub incoming_from: Vec<String>,
    pub size_bytes: u64,
    pub start_sector: u32,
    pub left_sibling_raw: u32,
    pub right_sibling_raw: u32,
    pub child_raw: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left_sibling: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right_sibling: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub child: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub created_filetime: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub modified_filetime: Option<u64>,
}

/// Inspect directory entries from a legacy CFB/OLE file on disk.
pub fn inspect_directory_path<P: AsRef<Path>>(
    path: P,
    config: &ParserConfig,
) -> AppResult<DirectoryInspection> {
    let path = path.as_ref();
    let metadata = fs::metadata(path).map_err(ParserParseError::from)?;
    if metadata.len() > config.max_input_size {
        return Err(ParserParseError::ResourceLimit(format!(
            "Input exceeds max_input_size ({} > {})",
            metadata.len(),
            config.max_input_size
        ))
        .into());
    }
    let bytes = fs::read(path).map_err(ParserParseError::from)?;
    inspect_directory_bytes(&bytes)
}

/// Inspect directory entries from raw CFB/OLE bytes.
pub fn inspect_directory_bytes(data: &[u8]) -> AppResult<DirectoryInspection> {
    let cfb = Cfb::parse(data.to_vec())?;
    let slot_count = cfb.list_directory_slots().len() as u32;
    let sector_count = cfb.sector_count();
    let mut entries: Vec<DirectoryEntry> = cfb
        .list_directory_slots()
        .into_iter()
        .map(|entry| DirectoryEntry {
            entry_index: entry.entry_index,
            path: entry.path.clone(),
            entry_type: entry
                .entry_type
                .map(map_entry_type)
                .unwrap_or("unknown")
                .to_string(),
            name_len_raw: entry.name_len_raw,
            object_type_raw: entry.object_type_raw,
            color_flag_raw: entry.color_flag_raw,
            state: map_state(entry.state).to_string(),
            classification: entry
                .entry_type
                .map(|entry_type| classify_entry(&entry.path, entry_type))
                .unwrap_or_else(|| "free-slot".to_string()),
            anomaly_severity: "none".to_string(),
            anomaly_tags: build_entry_anomaly_tags(EntryAnomalyContext {
                state: map_state(entry.state),
                entry_type: entry.entry_type.map(map_entry_type).unwrap_or("unknown"),
                name_len_raw: entry.name_len_raw,
                entry_index: entry.entry_index,
                slot_count,
                sector_count,
                start_sector: entry.start_sector,
                size_bytes: entry.size,
                left_sibling: entry.left_sibling,
                right_sibling: entry.right_sibling,
                child: entry.child,
            }),
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
            size_bytes: entry.size,
            start_sector: entry.start_sector,
            left_sibling_raw: entry.left_sibling_raw,
            right_sibling_raw: entry.right_sibling_raw,
            child_raw: entry.child_raw,
            left_sibling: entry.left_sibling,
            right_sibling: entry.right_sibling,
            child: entry.child,
            created_filetime: entry.created_filetime,
            modified_filetime: entry.modified_filetime,
        })
        .collect();
    annotate_incoming_references(&mut entries);
    entries.sort_by(|left, right| left.path.cmp(&right.path));
    let role_counts = build_role_counts(&entries);
    let anomaly_counts = build_anomaly_counts(&entries);
    let anomaly_catalog = build_anomaly_catalog(&entries);
    let anomaly_severity_counts = build_anomaly_severity_counts(&entries);
    let directory_score = build_directory_score(&entries);
    let reference_counts = build_reference_counts(&entries);
    let pointer_counts = build_pointer_counts(&entries);
    let tree_density_counts = build_tree_density_counts(&entries);
    let dangling_state_counts = build_dangling_state_counts(&entries);
    let self_reference_counts = build_self_reference_counts(&entries);
    let short_cycle_counts = build_short_cycle_counts(&entries);
    let reachability_counts = build_reachability_counts(&entries);
    let incoming_source_counts = build_incoming_source_counts(&entries);
    let incoming_source_type_counts = build_incoming_source_type_counts(&entries);
    let dead_reference_counts = build_dead_reference_counts(&entries);
    let fanout_counts = build_fanout_counts(&entries);
    Ok(DirectoryInspection {
        container: "cfb-ole".to_string(),
        entry_count: entries.len(),
        directory_score,
        role_counts,
        anomaly_counts,
        anomaly_catalog,
        anomaly_severity_counts,
        reference_counts,
        pointer_counts,
        tree_density_counts,
        dangling_state_counts,
        self_reference_counts,
        short_cycle_counts,
        reachability_counts,
        incoming_source_counts,
        incoming_source_type_counts,
        dead_reference_counts,
        fanout_counts,
        entries,
    })
}

fn map_state(state: CfbDirectoryState) -> &'static str {
    match state {
        CfbDirectoryState::Normal => "normal",
        CfbDirectoryState::Free => "free",
        CfbDirectoryState::Orphaned => "orphaned",
    }
}

fn map_entry_type(entry_type: CfbEntryType) -> &'static str {
    match entry_type {
        CfbEntryType::RootStorage => "root-storage",
        CfbEntryType::Storage => "storage",
        CfbEntryType::Stream => "stream",
    }
}

fn classify_entry(path: &str, entry_type: CfbEntryType) -> String {
    let upper = path.to_ascii_uppercase();
    match entry_type {
        CfbEntryType::RootStorage => "root-storage".to_string(),
        CfbEntryType::Storage => {
            if upper == "OBJECTPOOL" || upper.starts_with("OBJECTPOOL/") {
                "embedded-object-storage".to_string()
            } else if upper == "VBA" || upper.ends_with("/VBA") {
                "vba-storage".to_string()
            } else {
                "storage".to_string()
            }
        }
        CfbEntryType::Stream => {
            if upper == "WORDDOCUMENT" {
                "word-main-stream".to_string()
            } else if upper == "WORKBOOK" || upper == "BOOK" {
                "excel-main-stream".to_string()
            } else if upper == "POWERPOINT DOCUMENT" {
                "powerpoint-main-stream".to_string()
            } else if upper.ends_with("/PROJECT") || upper == "PROJECT" {
                "vba-project-metadata".to_string()
            } else if upper.contains("/VBA/") || upper.starts_with("VBA/") {
                "vba-module-stream".to_string()
            } else if upper.ends_with("OLE10NATIVE") {
                "ole-native-payload".to_string()
            } else if upper == "PACKAGE" || upper.ends_with("/PACKAGE") {
                "package-payload".to_string()
            } else if upper.ends_with("/CONTENTS") {
                "embedded-contents".to_string()
            } else {
                "stream".to_string()
            }
        }
    }
}

fn build_role_counts(entries: &[DirectoryEntry]) -> Vec<DirectoryRoleCount> {
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
        .map(|(role, count)| DirectoryRoleCount { role, count })
        .collect()
}

fn build_anomaly_counts(entries: &[DirectoryEntry]) -> Vec<DirectoryAnomalyCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        for anomaly in &entry.anomaly_tags {
            *counts.entry(anomaly.clone()).or_insert(0) += 1;
        }
    }
    counts
        .into_iter()
        .map(|(anomaly, count)| DirectoryAnomalyCount { anomaly, count })
        .collect()
}

fn anomaly_severity(tag: &str) -> &'static str {
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

fn build_anomaly_catalog(entries: &[DirectoryEntry]) -> Vec<DirectoryAnomalySeverity> {
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

fn build_anomaly_severity_counts(entries: &[DirectoryEntry]) -> Vec<DirectoryAnomalySeverityCount> {
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
        .map(|(severity, count)| DirectoryAnomalySeverityCount { severity, count })
        .collect()
}

fn severity_rank(value: &str) -> u8 {
    match value {
        "high" => 3,
        "medium" => 2,
        "low" => 1,
        _ => 0,
    }
}

fn severity_label(rank: u8) -> &'static str {
    match rank {
        3 => "high",
        2 => "medium",
        1 => "low",
        _ => "none",
    }
}

fn max_severity<I>(values: I) -> String
where
    I: IntoIterator<Item = String>,
{
    let mut best_rank = 0;
    for value in values {
        let rank = severity_rank(&value);
        if rank > best_rank {
            best_rank = rank;
        }
    }
    severity_label(best_rank).to_string()
}

fn build_directory_score(entries: &[DirectoryEntry]) -> String {
    max_severity(
        entries
            .iter()
            .filter(|entry| entry.state != "free")
            .map(|entry| entry.anomaly_severity.clone()),
    )
}

fn build_reference_counts(entries: &[DirectoryEntry]) -> Vec<DirectoryReferenceCount> {
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
        .map(|(bucket, count)| DirectoryReferenceCount { bucket, count })
        .collect()
}

fn build_pointer_counts(entries: &[DirectoryEntry]) -> Vec<DirectoryPointerCount> {
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
        .map(|(bucket, count)| DirectoryPointerCount { bucket, count })
        .collect()
}

fn build_tree_density_counts(entries: &[DirectoryEntry]) -> Vec<DirectoryTreeDensityCount> {
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
        .map(|(bucket, count)| DirectoryTreeDensityCount { bucket, count })
        .collect()
}

fn build_dangling_state_counts(entries: &[DirectoryEntry]) -> Vec<DirectoryPointerStateCount> {
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
        .map(|(bucket, count)| DirectoryPointerStateCount { bucket, count })
        .collect()
}

fn build_self_reference_counts(entries: &[DirectoryEntry]) -> Vec<DirectorySelfReferenceCount> {
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
        .map(|(bucket, count)| DirectorySelfReferenceCount { bucket, count })
        .collect()
}

fn build_short_cycle_counts(entries: &[DirectoryEntry]) -> Vec<DirectoryCycleCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in entries {
        for bucket in &entry.short_cycles {
            *counts.entry(bucket.clone()).or_insert(0) += 1;
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| DirectoryCycleCount { bucket, count })
        .collect()
}

fn build_reachability_counts(entries: &[DirectoryEntry]) -> Vec<DirectoryReachabilityCount> {
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
        .map(|(bucket, count)| DirectoryReachabilityCount { bucket, count })
        .collect()
}

fn build_incoming_source_counts(entries: &[DirectoryEntry]) -> Vec<DirectoryIncomingSourceCount> {
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
        .map(|(bucket, count)| DirectoryIncomingSourceCount { bucket, count })
        .collect()
}

fn build_incoming_source_type_counts(
    entries: &[DirectoryEntry],
) -> Vec<DirectoryIncomingSourceTypeCount> {
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
        .map(|(bucket, count)| DirectoryIncomingSourceTypeCount { bucket, count })
        .collect()
}

fn build_dead_reference_counts(entries: &[DirectoryEntry]) -> Vec<DirectoryDeadReferenceCount> {
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
        .map(|(bucket, count)| DirectoryDeadReferenceCount { bucket, count })
        .collect()
}

fn build_fanout_counts(entries: &[DirectoryEntry]) -> Vec<DirectoryFanoutCount> {
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
        .map(|(bucket, count)| DirectoryFanoutCount { bucket, count })
        .collect()
}

fn build_entry_anomaly_tags(context: EntryAnomalyContext<'_>) -> Vec<String> {
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

fn annotate_incoming_references(entries: &mut [DirectoryEntry]) {
    let path_by_index = entries
        .iter()
        .map(|entry| {
            (
                entry.entry_index,
                (
                    entry.path.clone(),
                    entry.state.clone(),
                    entry.entry_type.clone(),
                    entry.left_sibling,
                    entry.right_sibling,
                    entry.child,
                ),
            )
        })
        .collect::<std::collections::BTreeMap<_, _>>();
    let mut incoming = std::collections::BTreeMap::<u32, Vec<(String, bool, String)>>::new();
    let reachable = collect_reachable_indices(entries);

    let mut short_cycles_by_index = std::collections::BTreeMap::<u32, Vec<String>>::new();
    for entry in entries.iter() {
        if entry.state == "free" {
            continue;
        }
        for (label, pointer) in [
            ("left", entry.left_sibling),
            ("right", entry.right_sibling),
            ("child", entry.child),
        ] {
            if let Some(target) = pointer {
                if let Some((_, source_state, source_entry_type, _, _, _)) =
                    path_by_index.get(&entry.entry_index)
                {
                    incoming.entry(target).or_default().push((
                        format!("{label}:{}#{}", entry.path, entry.entry_index),
                        source_state == "normal",
                        source_entry_type.clone(),
                    ));
                }
                if let Some((_, _, _, target_left, target_right, target_child)) =
                    path_by_index.get(&target)
                {
                    let reverse_hits = [
                        ("left", *target_left == Some(entry.entry_index)),
                        ("right", *target_right == Some(entry.entry_index)),
                        ("child", *target_child == Some(entry.entry_index)),
                    ];
                    let same_bucket = match label {
                        "child" if *target_child == Some(entry.entry_index) => {
                            Some("child-2-cycle")
                        }
                        "left" | "right"
                            if *target_left == Some(entry.entry_index)
                                || *target_right == Some(entry.entry_index) =>
                        {
                            Some("sibling-2-cycle")
                        }
                        _ => None,
                    };
                    if let Some(cycle_bucket) = same_bucket {
                        short_cycles_by_index
                            .entry(entry.entry_index)
                            .or_default()
                            .push(cycle_bucket.to_string());
                    }
                    if reverse_hits
                        .iter()
                        .any(|(target_label, hit)| *hit && *target_label != label)
                    {
                        short_cycles_by_index
                            .entry(entry.entry_index)
                            .or_default()
                            .push("mixed-2-cycle".to_string());
                    }
                }
                if let Some((_, _, _, target_left, target_right, target_child)) =
                    path_by_index.get(&target)
                {
                    for (target_label, next_opt) in [
                        ("left", *target_left),
                        ("right", *target_right),
                        ("child", *target_child),
                    ] {
                        let Some(next) = next_opt else { continue };
                        if let Some((_, _, _, next_left, next_right, next_child)) =
                            path_by_index.get(&next)
                        {
                            let closing_hits = [
                                ("left", *next_left == Some(entry.entry_index)),
                                ("right", *next_right == Some(entry.entry_index)),
                                ("child", *next_child == Some(entry.entry_index)),
                            ];
                            let same_bucket = match label {
                                "child" if closing_hits.iter().any(|(_, hit)| *hit) => {
                                    Some("child-3-cycle")
                                }
                                "left" | "right"
                                    if closing_hits.iter().any(|(next_label, hit)| {
                                        *hit && (*next_label == "left" || *next_label == "right")
                                    }) =>
                                {
                                    Some("sibling-3-cycle")
                                }
                                _ => None,
                            };
                            if let Some(cycle_bucket) = same_bucket {
                                short_cycles_by_index
                                    .entry(entry.entry_index)
                                    .or_default()
                                    .push(cycle_bucket.to_string());
                            }
                            if closing_hits.iter().any(|(_, hit)| *hit)
                                && ([label, target_label]
                                    .into_iter()
                                    .chain(
                                        closing_hits
                                            .iter()
                                            .filter(|(_, hit)| *hit)
                                            .map(|(next_label, _)| *next_label),
                                    )
                                    .collect::<std::collections::BTreeSet<_>>()
                                    .len()
                                    > 1)
                            {
                                short_cycles_by_index
                                    .entry(entry.entry_index)
                                    .or_default()
                                    .push("mixed-3-cycle".to_string());
                            }
                        }
                    }
                }
            }
        }
    }

    for entry in entries.iter_mut() {
        let refs = incoming.remove(&entry.entry_index).unwrap_or_default();
        entry.fanout_count = [entry.left_sibling, entry.right_sibling, entry.child]
            .into_iter()
            .flatten()
            .count();
        let cycles = short_cycles_by_index
            .remove(&entry.entry_index)
            .unwrap_or_default();
        for cycle in cycles {
            if !entry.short_cycles.iter().any(|item| item == &cycle) {
                entry.short_cycles.push(cycle.clone());
            }
            if !entry.anomaly_tags.iter().any(|item| item == &cycle) {
                entry.anomaly_tags.push(cycle);
            }
        }
        entry.incoming_reference_count = refs.len();
        entry.incoming_normal_reference_count =
            refs.iter().filter(|(_, normal, _)| *normal).count();
        entry.incoming_anomalous_reference_count =
            refs.len() - entry.incoming_normal_reference_count;
        entry.incoming_from_root_storage_count = refs
            .iter()
            .filter(|(_, _, source_type)| source_type == "root-storage")
            .count();
        entry.incoming_from_storage_count = refs
            .iter()
            .filter(|(_, _, source_type)| source_type == "storage")
            .count();
        entry.incoming_from_stream_count = refs
            .iter()
            .filter(|(_, _, source_type)| source_type == "stream")
            .count();
        entry.incoming_from = refs.into_iter().map(|(label, _, _)| label).collect();
        entry.reachable_from_root = reachable.contains(&entry.entry_index);

        if entry.state == "normal"
            && entry.entry_type != "root-storage"
            && entry.incoming_reference_count == 0
        {
            entry
                .anomaly_tags
                .push("unreferenced-live-entry".to_string());
        }
        if entry.incoming_reference_count > 1 {
            entry
                .anomaly_tags
                .push("multi-referenced-entry".to_string());
        }
        if entry.state == "normal"
            && entry.entry_type != "root-storage"
            && !entry.reachable_from_root
        {
            entry
                .anomaly_tags
                .push("unreachable-live-entry".to_string());
        }
        if entry.incoming_anomalous_reference_count > 0 {
            entry
                .anomaly_tags
                .push("incoming-from-anomalous-entry".to_string());
        }
        entry.anomaly_severity = max_severity(
            entry
                .anomaly_tags
                .iter()
                .map(|tag| anomaly_severity(tag).to_string()),
        );
    }
}

fn collect_reachable_indices(entries: &[DirectoryEntry]) -> std::collections::HashSet<u32> {
    let by_index = entries
        .iter()
        .map(|entry| {
            (
                entry.entry_index,
                (entry.left_sibling, entry.right_sibling, entry.child),
            )
        })
        .collect::<std::collections::BTreeMap<_, _>>();
    let mut reachable = std::collections::HashSet::new();
    let mut queue = std::collections::VecDeque::new();
    queue.push_back(0u32);

    while let Some(index) = queue.pop_front() {
        if !reachable.insert(index) {
            continue;
        }
        if let Some((left, right, child)) = by_index.get(&index) {
            for next in [*left, *right, *child].into_iter().flatten() {
                if by_index.contains_key(&next) && !reachable.contains(&next) {
                    queue.push_back(next);
                }
            }
        }
    }

    reachable
}

#[cfg(test)]
mod tests {
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
            .any(|entry| entry.role == "state:normal" && entry.count >= 1));
        assert!(inspection
            .role_counts
            .iter()
            .any(|entry| entry.role == "classification:word-main-stream" && entry.count == 1));
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
            .any(|entry| entry.anomaly == "free-slot" && entry.count == 1));
        assert!(anomalies
            .iter()
            .any(|entry| entry.anomaly == "orphaned-entry" && entry.count == 1));
        assert!(anomalies
            .iter()
            .any(|entry| entry.anomaly == "unknown-entry-type" && entry.count == 1));
        assert!(anomalies
            .iter()
            .any(|entry| entry.anomaly == "zero-name-length" && entry.count == 1));
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
            .any(|entry| entry.severity == "high" && entry.count >= 1));
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
}
