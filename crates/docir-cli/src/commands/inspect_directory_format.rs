use docir_app::DirectoryInspection;

use crate::commands::util::{push_bullet_line, push_count_section, push_labeled_line};

pub(super) fn format_inspection_text(inspection: &DirectoryInspection) -> String {
    let mut out = String::new();
    format_header_and_summary(&mut out, inspection);
    format_entries_section(&mut out, &inspection.entries);
    out
}

fn format_header_and_summary(out: &mut String, inspection: &DirectoryInspection) {
    push_labeled_line(out, 0, "Container", &inspection.container);
    push_labeled_line(out, 0, "Entries", inspection.entry_count);
    push_labeled_line(out, 0, "Directory Score", &inspection.directory_score);
    push_count_section(
        out,
        "Role Counts",
        &inspection.role_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        out,
        "Anomalies",
        &inspection.anomaly_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    if !inspection.anomaly_catalog.is_empty() {
        out.push_str("\nAnomaly Severity Catalog:\n");
        for entry in &inspection.anomaly_catalog {
            push_bullet_line(out, 2, &entry.anomaly, &entry.severity);
        }
    }
    format_anomaly_summary(out, inspection);
    format_reference_summary(out, inspection);
    format_fanout_summary(out, inspection);
}

fn format_anomaly_summary(out: &mut String, inspection: &DirectoryInspection) {
    push_count_section(
        out,
        "Anomaly Severity Summary",
        &inspection.anomaly_severity_counts,
        |e| &e.bucket,
        |e| e.count,
    );
}

fn format_reference_summary(out: &mut String, inspection: &DirectoryInspection) {
    push_count_section(
        out,
        "Reference Summary",
        &inspection.reference_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        out,
        "Pointer Summary",
        &inspection.pointer_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        out,
        "Tree Density Summary",
        &inspection.tree_density_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        out,
        "Dangling By State",
        &inspection.dangling_state_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        out,
        "Self References",
        &inspection.self_reference_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        out,
        "Short Cycles",
        &inspection.short_cycle_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        out,
        "Reachability Summary",
        &inspection.reachability_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        out,
        "Incoming Source Summary",
        &inspection.incoming_source_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        out,
        "Incoming Source Types",
        &inspection.incoming_source_type_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        out,
        "Dead But Referenced",
        &inspection.dead_reference_counts,
        |e| &e.bucket,
        |e| e.count,
    );
}

fn format_fanout_summary(out: &mut String, inspection: &DirectoryInspection) {
    push_count_section(
        out,
        "Fanout Summary",
        &inspection.fanout_counts,
        |e| &e.bucket,
        |e| e.count,
    );
}

fn format_entries_section(out: &mut String, entries: &[docir_app::DirectoryEntry]) {
    if entries.is_empty() {
        return;
    }
    out.push_str("\nDirectory:\n");
    for entry in entries {
        format_directory_entry(out, entry);
    }
}

fn format_directory_entry(out: &mut String, entry: &docir_app::DirectoryEntry) {
    push_bullet_line(out, 2, &entry.entry_type, &entry.path);
    push_labeled_line(out, 4, "Entry Index", entry.entry_index);
    push_labeled_line(out, 4, "Name Length Raw", entry.name_len_raw);
    push_labeled_line(out, 4, "Object Type Raw", entry.object_type_raw);
    push_labeled_line(out, 4, "Color Flag Raw", entry.color_flag_raw);
    push_labeled_line(out, 4, "State", &entry.state);
    push_labeled_line(out, 4, "Classification", &entry.classification);
    push_labeled_line(out, 4, "Anomaly Severity", &entry.anomaly_severity);
    if !entry.anomaly_tags.is_empty() {
        push_labeled_line(out, 4, "Anomaly Tags", entry.anomaly_tags.join(", "));
    }
    if !entry.short_cycles.is_empty() {
        push_labeled_line(out, 4, "Short Cycles", entry.short_cycles.join(", "));
    }
    push_labeled_line(out, 4, "Reachable From Root", entry.reachable_from_root);
    push_labeled_line(
        out,
        4,
        "Incoming References",
        entry.incoming_reference_count,
    );
    push_labeled_line(
        out,
        4,
        "Incoming From Normal",
        entry.incoming_normal_reference_count,
    );
    push_labeled_line(
        out,
        4,
        "Incoming From Anomalous",
        entry.incoming_anomalous_reference_count,
    );
    push_labeled_line(
        out,
        4,
        "Incoming From Root Storage",
        entry.incoming_from_root_storage_count,
    );
    push_labeled_line(
        out,
        4,
        "Incoming From Storage",
        entry.incoming_from_storage_count,
    );
    push_labeled_line(
        out,
        4,
        "Incoming From Stream",
        entry.incoming_from_stream_count,
    );
    push_labeled_line(out, 4, "Fanout", entry.fanout_count);
    if !entry.incoming_from.is_empty() {
        push_labeled_line(out, 4, "Incoming From", entry.incoming_from.join(", "));
    }
    push_labeled_line(out, 4, "Size", format!("{} bytes", entry.size_bytes));
    push_labeled_line(out, 4, "Sector", entry.start_sector);
    push_labeled_line(out, 4, "Left Sibling Raw", entry.left_sibling_raw);
    push_labeled_line(out, 4, "Right Sibling Raw", entry.right_sibling_raw);
    push_labeled_line(out, 4, "Child Raw", entry.child_raw);
    if let Some(left) = entry.left_sibling {
        push_labeled_line(out, 4, "Left Sibling", left);
    }
    if let Some(right) = entry.right_sibling {
        push_labeled_line(out, 4, "Right Sibling", right);
    }
    if let Some(child) = entry.child {
        push_labeled_line(out, 4, "Child", child);
    }
    if let Some(created) = entry.created_filetime {
        push_labeled_line(out, 4, "Created", created);
    }
    if let Some(modified) = entry.modified_filetime {
        push_labeled_line(out, 4, "Modified", modified);
    }
}
