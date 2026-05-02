mod anomaly;
mod classify;
mod references;
mod statistics;
mod types;

#[cfg(test)]
mod tests;

use crate::io_support::read_bounded_file;
use crate::{AppResult, ParserConfig};
use std::path::Path;

pub use types::{DirectoryAnomalySeverity, DirectoryEntry, DirectoryInspection};

use anomaly::{
    build_anomaly_catalog, build_anomaly_counts, build_anomaly_severity_counts,
    build_directory_score, build_entry_anomaly_tags, build_role_counts,
};
use classify::{classify_entry, map_entry_type, map_state};
use references::annotate_incoming_references;
use statistics::{
    build_dangling_state_counts, build_dead_reference_counts, build_fanout_counts,
    build_incoming_source_counts, build_incoming_source_type_counts, build_pointer_counts,
    build_reachability_counts, build_reference_counts, build_self_reference_counts,
    build_short_cycle_counts, build_tree_density_counts,
};
use types::EntryAnomalyContext;

/// Inspect directory entries from a legacy CFB/OLE file on disk.
pub fn inspect_directory_path<P: AsRef<Path>>(
    path: P,
    config: &ParserConfig,
) -> AppResult<DirectoryInspection> {
    let bytes = read_bounded_file(path, config.max_input_size)?;
    inspect_directory_bytes(&bytes)
}

/// Inspect directory entries from raw CFB/OLE bytes.
pub fn inspect_directory_bytes(data: &[u8]) -> AppResult<DirectoryInspection> {
    let cfb = docir_parser::ole::Cfb::parse(data.to_vec())?;
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
