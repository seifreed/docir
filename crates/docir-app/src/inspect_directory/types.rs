use crate::bucket_count::BucketCount;
use serde::Serialize;

/// Dedicated structural view of normal CFB directory entries.
#[derive(Debug, Clone, Serialize)]
pub struct DirectoryInspection {
    pub container: String,
    pub entry_count: usize,
    pub directory_score: String,
    pub role_counts: Vec<BucketCount>,
    pub anomaly_counts: Vec<BucketCount>,
    pub anomaly_catalog: Vec<DirectoryAnomalySeverity>,
    pub anomaly_severity_counts: Vec<BucketCount>,
    pub reference_counts: Vec<BucketCount>,
    pub pointer_counts: Vec<BucketCount>,
    pub tree_density_counts: Vec<BucketCount>,
    pub dangling_state_counts: Vec<BucketCount>,
    pub self_reference_counts: Vec<BucketCount>,
    pub short_cycle_counts: Vec<BucketCount>,
    pub reachability_counts: Vec<BucketCount>,
    pub incoming_source_counts: Vec<BucketCount>,
    pub incoming_source_type_counts: Vec<BucketCount>,
    pub dead_reference_counts: Vec<BucketCount>,
    pub fanout_counts: Vec<BucketCount>,
    pub entries: Vec<DirectoryEntry>,
}

/// An anomaly entry with its severity classification.
#[derive(Debug, Clone, Serialize)]
pub struct DirectoryAnomalySeverity {
    pub anomaly: String,
    pub severity: String,
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct EntryAnomalyContext<'a> {
    pub state: &'a str,
    pub entry_type: &'a str,
    pub name_len_raw: u16,
    pub entry_index: u32,
    pub slot_count: u32,
    pub sector_count: u32,
    pub start_sector: u32,
    pub size_bytes: u64,
    pub left_sibling: Option<u32>,
    pub right_sibling: Option<u32>,
    pub child: Option<u32>,
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