use docir_app::SectorInspection;

use crate::commands::util::{push_bullet_line, push_count_section, push_labeled_line};

pub(super) fn format_inspection_text(inspection: &SectorInspection) -> String {
    let mut out = String::new();
    format_header_fields(&mut out, inspection);
    format_list_sections(&mut out, inspection);
    format_sector_overview_section(&mut out, inspection);
    format_stream_chains_section(&mut out, inspection);
    format_anomalies_section(&mut out, &inspection.anomalies);
    out
}

fn format_header_fields(out: &mut String, inspection: &SectorInspection) {
    push_labeled_line(out, 0, "Container", &inspection.container);
    push_labeled_line(out, 0, "Sector Score", &inspection.sector_score);
    push_labeled_line(out, 0, "Sector Size", inspection.sector_size);
    push_labeled_line(out, 0, "Mini Sector Size", inspection.mini_sector_size);
    push_labeled_line(out, 0, "Mini Cutoff", inspection.mini_cutoff);
    push_labeled_line(out, 0, "Header FAT Sectors", inspection.num_fat_sectors);
    push_labeled_line(
        out,
        0,
        "Header First Directory Sector",
        inspection.first_dir_sector,
    );
    push_labeled_line(out, 0, "Header First MiniFAT", inspection.first_mini_fat);
    push_labeled_line(out, 0, "Header MiniFAT Sectors", inspection.num_mini_fat);
    push_labeled_line(out, 0, "Header First DIFAT", inspection.first_difat);
    push_labeled_line(out, 0, "Header DIFAT Sectors", inspection.num_difat);
    push_labeled_line(out, 0, "DIFAT Entries", inspection.difat_entry_count);
    push_labeled_line(out, 0, "Sectors", inspection.sector_count);
    push_labeled_line(out, 0, "FAT Entries", inspection.fat_entry_count);
    push_labeled_line(out, 0, "FAT Free", inspection.fat_free_count);
    push_labeled_line(out, 0, "FAT Occupied", inspection.occupied_fat_entries);
    push_labeled_line(
        out,
        0,
        "FAT End Of Chain",
        inspection.fat_end_of_chain_count,
    );
    push_labeled_line(out, 0, "FAT Reserved", inspection.fat_reserved_count);
    push_labeled_line(out, 0, "MiniFAT Entries", inspection.mini_fat_entry_count);
}

fn format_list_sections(out: &mut String, inspection: &SectorInspection) {
    push_count_section(
        out,
        "Role Counts",
        &inspection.role_counts,
        |e| &e.role,
        |e| e.count,
    );
    if !inspection.shared_sector_claims.is_empty() {
        out.push_str("\nShared Sectors:\n");
        for claim in &inspection.shared_sector_claims {
            push_bullet_line(out, 2, &claim.sector.to_string(), claim.owners.join(", "));
        }
    }
    if !inspection.shared_chain_overlaps.is_empty() {
        out.push_str("\nShared Chains:\n");
        for overlap in &inspection.shared_chain_overlaps {
            push_bullet_line(
                out,
                2,
                &format!("{} [{}]", overlap.owners.join(", "), overlap.severity),
                format!("{:?}", overlap.sectors),
            );
        }
    }
    if !inspection.start_sector_reuse.is_empty() {
        out.push_str("\nStart Sector Reuse:\n");
        for reuse in &inspection.start_sector_reuse {
            push_bullet_line(out, 2, &reuse.sector.to_string(), reuse.owners.join(", "));
        }
    }
    push_count_section(
        out,
        "Truncated Chains",
        &inspection.truncated_chain_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        out,
        "Chain Health By Root",
        &inspection.chain_health_by_root,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        out,
        "Chain Health By Allocation",
        &inspection.chain_health_by_allocation,
        |e| &e.bucket,
        |e| e.count,
    );
    if !inspection.structural_incoherence_counts.is_empty() {
        out.push_str("\nStructural Incoherence:\n");
        for entry in &inspection.structural_incoherence_counts {
            push_bullet_line(
                out,
                2,
                &entry.bucket,
                format!("{} [{}]", entry.count, entry.severity),
            );
        }
    }
}

fn format_sector_overview_section(out: &mut String, inspection: &SectorInspection) {
    if inspection.sector_overview.is_empty() {
        return;
    }
    out.push_str("\nSector Overview:\n");
    for sector in &inspection.sector_overview {
        push_bullet_line(out, 2, &sector.sector.to_string(), &sector.role);
        push_labeled_line(out, 4, "FAT Raw", sector.fat_raw);
        push_labeled_line(out, 4, "FAT Value", &sector.fat_value);
        if !sector.special_roles.is_empty() {
            push_labeled_line(out, 4, "Special Roles", sector.special_roles.join(", "));
        }
        if !sector.owners.is_empty() {
            let owners = sector
                .owners
                .iter()
                .map(|owner| {
                    format!(
                        "{}[{}]{}",
                        owner.path,
                        owner.index_in_chain,
                        if owner.is_terminal { " terminal" } else { "" }
                    )
                })
                .collect::<Vec<_>>()
                .join(", ");
            push_labeled_line(out, 4, "Owners", owners);
        }
    }
}

fn format_stream_chains_section(out: &mut String, inspection: &SectorInspection) {
    if inspection.streams.is_empty() {
        return;
    }
    out.push_str("\nStream Chains:\n");
    for stream in &inspection.streams {
        push_bullet_line(out, 2, &stream.path, &stream.allocation);
        push_labeled_line(out, 4, "Logical Root", &stream.logical_root);
        push_labeled_line(out, 4, "Chain State", &stream.chain_state);
        push_labeled_line(out, 4, "Stream Health", &stream.stream_health);
        push_labeled_line(out, 4, "Stream Risk", &stream.stream_risk);
        push_labeled_line(out, 4, "Size", format!("{} bytes", stream.size_bytes));
        push_labeled_line(out, 4, "Start Sector", stream.start_sector);
        push_labeled_line(out, 4, "Expected Chain", stream.expected_chain_len);
        push_labeled_line(out, 4, "Chain Terminal", &stream.chain_terminal);
        push_labeled_line(out, 4, "Chain Terminal Raw", stream.chain_terminal_raw);
        push_labeled_line(out, 4, "Chain", format!("{:?}", stream.sector_chain));
        if !stream.chain_steps.is_empty() {
            out.push_str("    Steps:\n");
            for step in &stream.chain_steps {
                push_bullet_line(
                    out,
                    6,
                    &step.sector.to_string(),
                    format!("{} ({})", step.next_raw, step.next),
                );
            }
        }
    }
}

fn format_anomalies_section(out: &mut String, anomalies: &[docir_app::SectorAnomaly]) {
    if anomalies.is_empty() {
        return;
    }
    out.push_str("\nAnomalies:\n");
    for anomaly in anomalies {
        push_bullet_line(
            out,
            2,
            &format!("{} [{}]", anomaly.kind, anomaly.severity),
            &anomaly.message,
        );
    }
}
