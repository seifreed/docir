use std::collections::{BTreeMap, BTreeSet, HashSet, VecDeque};

use super::types::DirectoryEntry;

type PathInfo = (
    String,
    String,
    String,
    Option<u32>,
    Option<u32>,
    Option<u32>,
);

pub(super) fn build_path_by_index(entries: &[DirectoryEntry]) -> BTreeMap<u32, PathInfo> {
    entries
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
        .collect()
}

pub(super) fn collect_incoming_refs(
    entries: &[DirectoryEntry],
    path_by_index: &BTreeMap<u32, PathInfo>,
) -> BTreeMap<u32, Vec<(String, bool, String)>> {
    let mut incoming = BTreeMap::<u32, Vec<(String, bool, String)>>::new();
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
            }
        }
    }
    incoming
}

pub(super) fn detect_2_cycles(
    entry: &DirectoryEntry,
    path_by_index: &BTreeMap<u32, PathInfo>,
    label: &str,
    target: u32,
) -> Vec<String> {
    let mut cycles = Vec::new();
    if let Some((_, _, _, target_left, target_right, target_child)) = path_by_index.get(&target) {
        let reverse_hits = [
            ("left", *target_left == Some(entry.entry_index)),
            ("right", *target_right == Some(entry.entry_index)),
            ("child", *target_child == Some(entry.entry_index)),
        ];
        let same_bucket = match label {
            "child" if *target_child == Some(entry.entry_index) => Some("child-2-cycle"),
            "left" | "right"
                if *target_left == Some(entry.entry_index)
                    || *target_right == Some(entry.entry_index) =>
            {
                Some("sibling-2-cycle")
            }
            _ => None,
        };
        if let Some(cycle_bucket) = same_bucket {
            cycles.push(cycle_bucket.to_string());
        }
        if reverse_hits
            .iter()
            .any(|(target_label, hit)| *hit && *target_label != label)
        {
            cycles.push("mixed-2-cycle".to_string());
        }
    }
    cycles
}

pub(super) fn detect_3_cycles(
    entry: &DirectoryEntry,
    path_by_index: &BTreeMap<u32, PathInfo>,
    label: &str,
    target: u32,
) -> Vec<String> {
    let mut cycles = Vec::new();
    if let Some((_, _, _, target_left, target_right, target_child)) = path_by_index.get(&target) {
        for (target_label, next_opt) in [
            ("left", *target_left),
            ("right", *target_right),
            ("child", *target_child),
        ] {
            let Some(next) = next_opt else { continue };
            if let Some((_, _, _, next_left, next_right, next_child)) = path_by_index.get(&next) {
                let closing_hits = [
                    ("left", *next_left == Some(entry.entry_index)),
                    ("right", *next_right == Some(entry.entry_index)),
                    ("child", *next_child == Some(entry.entry_index)),
                ];
                let same_bucket = match label {
                    "child" if closing_hits.iter().any(|(_, hit)| *hit) => Some("child-3-cycle"),
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
                    cycles.push(cycle_bucket.to_string());
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
                        .collect::<BTreeSet<_>>()
                        .len()
                        > 1)
                {
                    cycles.push("mixed-3-cycle".to_string());
                }
            }
        }
    }
    cycles
}

pub(super) fn collect_reachable_indices(entries: &[DirectoryEntry]) -> HashSet<u32> {
    let by_index = entries
        .iter()
        .map(|entry| {
            (
                entry.entry_index,
                (entry.left_sibling, entry.right_sibling, entry.child),
            )
        })
        .collect::<BTreeMap<_, _>>();
    let mut reachable = HashSet::new();
    let mut queue = VecDeque::new();
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

pub(super) fn apply_incoming_refs_to_entries(
    entries: &mut [DirectoryEntry],
    incoming: BTreeMap<u32, Vec<(String, bool, String)>>,
    cycles: BTreeMap<u32, Vec<String>>,
    reachable: &HashSet<u32>,
) {
    for entry in entries.iter_mut() {
        let refs = incoming
            .get(&entry.entry_index)
            .cloned()
            .unwrap_or_default();
        entry.fanout_count = count_fanout(entry);
        apply_short_cycles(entry, cycles.get(&entry.entry_index).map(Vec::as_slice));
        apply_incoming_reference_counts(entry, &refs);
        entry.incoming_from = refs.into_iter().map(|(label, _, _)| label).collect();
        entry.reachable_from_root = reachable.contains(&entry.entry_index);
        apply_reference_anomaly_tags(entry);
        refresh_entry_anomaly_severity(entry);
    }
}

fn count_fanout(entry: &DirectoryEntry) -> usize {
    [entry.left_sibling, entry.right_sibling, entry.child]
        .into_iter()
        .flatten()
        .count()
}

fn apply_short_cycles(entry: &mut DirectoryEntry, cycles: Option<&[String]>) {
    for cycle in cycles.into_iter().flatten() {
        if !entry.short_cycles.iter().any(|item| item == cycle) {
            entry.short_cycles.push(cycle.clone());
        }
        if !entry.anomaly_tags.iter().any(|item| item == cycle) {
            entry.anomaly_tags.push(cycle.clone());
        }
    }
}

fn apply_incoming_reference_counts(entry: &mut DirectoryEntry, refs: &[(String, bool, String)]) {
    entry.incoming_reference_count = refs.len();
    entry.incoming_normal_reference_count = refs.iter().filter(|(_, normal, _)| *normal).count();
    entry.incoming_anomalous_reference_count = refs.len() - entry.incoming_normal_reference_count;
    entry.incoming_from_root_storage_count = count_refs_from_type(refs, "root-storage");
    entry.incoming_from_storage_count = count_refs_from_type(refs, "storage");
    entry.incoming_from_stream_count = count_refs_from_type(refs, "stream");
}

fn count_refs_from_type(refs: &[(String, bool, String)], source_type: &str) -> usize {
    refs.iter()
        .filter(|(_, _, candidate)| candidate == source_type)
        .count()
}

fn apply_reference_anomaly_tags(entry: &mut DirectoryEntry) {
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
    if entry.state == "normal" && entry.entry_type != "root-storage" && !entry.reachable_from_root {
        entry
            .anomaly_tags
            .push("unreachable-live-entry".to_string());
    }
    if entry.incoming_anomalous_reference_count > 0 {
        entry
            .anomaly_tags
            .push("incoming-from-anomalous-entry".to_string());
    }
}

fn refresh_entry_anomaly_severity(entry: &mut DirectoryEntry) {
    use super::anomaly::anomaly_severity;
    use crate::severity::max_severity;

    entry.anomaly_severity = max_severity(
        entry
            .anomaly_tags
            .iter()
            .map(|tag| anomaly_severity(tag).to_string()),
    );
}

pub(super) fn annotate_incoming_references(entries: &mut [DirectoryEntry]) {
    let path_by_index = build_path_by_index(entries);
    let incoming = collect_incoming_refs(entries, &path_by_index);
    let reachable = collect_reachable_indices(entries);

    let mut short_cycles_by_index = BTreeMap::<u32, Vec<String>>::new();
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
                for cycle in detect_2_cycles(entry, &path_by_index, label, target) {
                    short_cycles_by_index
                        .entry(entry.entry_index)
                        .or_default()
                        .push(cycle);
                }
                for cycle in detect_3_cycles(entry, &path_by_index, label, target) {
                    short_cycles_by_index
                        .entry(entry.entry_index)
                        .or_default()
                        .push(cycle);
                }
            }
        }
    }

    apply_incoming_refs_to_entries(entries, incoming, short_cycles_by_index, &reachable);
}
