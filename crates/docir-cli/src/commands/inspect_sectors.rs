//! Inspect CFB sector allocation and stream chains.

use anyhow::Result;
use docir_app::{inspect_sectors_path, ParserConfig, SectorInspection};
use std::path::PathBuf;

use crate::cli::JsonOutputOpts;
use crate::commands::util::{
    push_bullet_line, push_count_section, push_labeled_line, run_dual_output,
};

/// Public API entrypoint: run.
pub fn run(input: PathBuf, opts: JsonOutputOpts, parser_config: &ParserConfig) -> Result<()> {
    let JsonOutputOpts {
        json,
        pretty,
        output,
    } = opts;
    let inspection = inspect_sectors_path(&input, parser_config)?;
    run_dual_output(
        &inspection,
        "inspection",
        json,
        pretty,
        output,
        format_inspection_text,
    )
}

fn format_inspection_text(inspection: &SectorInspection) -> String {
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

#[cfg(test)]
mod tests {
    use super::{format_inspection_text, run};
    use crate::cli::JsonOutputOpts;
    use crate::test_support;
    use docir_app::{
        test_support::build_test_cfb, ChainHealthCount, ChainStep, ParserConfig, RoleCount,
        SectorAnomaly, SectorInspection, SectorOverviewEntry, SharedChainOverlap,
        SharedSectorClaim, StartSectorReuse, StreamSectorMap, StructuralIncoherenceCount,
        TruncatedChainCount,
    };
    use std::fs;

    #[test]
    fn inspect_sectors_run_writes_json() {
        let input = test_support::temp_file("legacy", "doc");
        let output = test_support::temp_file("legacy", "json");
        fs::write(
            &input,
            build_test_cfb(&[("WordDocument", b"doc"), ("VBA/PROJECT", b"meta")]),
        )
        .expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: true,
                pretty: true,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("inspect-sectors json");

        let text = fs::read_to_string(&output).expect("output");
        assert!(text.contains("\"sector_size\": 512"));
        assert!(text.contains("\"sector_score\": "));
        assert!(text.contains("\"num_fat_sectors\": 1"));
        assert!(text.contains("\"path\": \"WordDocument\""));
        assert!(text.contains("\"logical_root\": \"WordDocument\""));
        assert!(text.contains("\"chain_state\": \"complete\""));
        assert!(text.contains("\"chain_terminal_raw\": 4294967294"));
        assert!(text.contains("\"chain_steps\": ["));
        assert!(text.contains("\"role_counts\": ["));
        assert!(text.contains("\"sector_overview\": ["));
        assert!(text.contains("\"special_roles\": ["));
        assert!(text.contains("\"owners\": ["));
        assert!(text.contains("\"index_in_chain\": 0"));
        assert!(text.contains("\"is_terminal\": true"));
        assert!(text.contains("\"start_sector\": 0"));
        assert!(text.contains("\"sector_chain\": ["));
        assert!(text.contains("\"occupied_fat_entries\":"));
        assert!(text.contains("\"shared_sector_claims\": ["));
        assert!(text.contains("\"shared_chain_overlaps\": ["));
        assert!(text.contains("\"start_sector_reuse\": ["));
        assert!(text.contains("\"truncated_chain_counts\": ["));
        assert!(text.contains("\"structural_incoherence_counts\": ["));
        assert!(text.contains("\"chain_health_by_root\": ["));
        assert!(text.contains("\"chain_health_by_allocation\": ["));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn inspect_sectors_run_writes_text() {
        let input = test_support::temp_file("legacy_text", "doc");
        let output = test_support::temp_file("legacy_text", "txt");
        fs::write(&input, build_test_cfb(&[("WordDocument", b"doc")])).expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: false,
                pretty: false,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("inspect-sectors text");

        let text = fs::read_to_string(&output).expect("output");
        assert!(text.contains("Sector Size: 512"));
        assert!(text.contains("Header FAT Sectors: 1"));
        assert!(text.contains("Header First Directory Sector:"));
        assert!(text.contains("FAT Occupied:"));
        assert!(text.contains("Role Counts:"));
        assert!(text.contains("stream-data:"));
        assert!(!text.contains("Structural Incoherence:"));
        assert!(text.contains("Sector Overview:"));
        assert!(text.contains("0: stream-data"));
        assert!(text.contains("Special Roles: directory"));
        assert!(text.contains("Owners: WordDocument[0] terminal"));
        assert!(text.contains("Stream Chains:"));
        assert!(text.contains("WordDocument: fat"));
        assert!(text.contains("Logical Root: WordDocument"));
        assert!(text.contains("Chain State: complete"));
        assert!(text.contains("Size: 3 bytes"));
        assert!(text.contains("Start Sector: 0"));
        assert!(text.contains("Expected Chain: 1"));
        assert!(text.contains("Chain Terminal: end-of-chain"));
        assert!(text.contains("Steps:"));
        assert!(text.contains("0: 4294967294 (end-of-chain)"));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn format_inspection_text_renders_expected_fields() {
        let inspection = SectorInspection {
            container: "cfb-ole".to_string(),
            sector_score: "high".to_string(),
            sector_size: 512,
            mini_sector_size: 64,
            mini_cutoff: 4096,
            num_fat_sectors: 1,
            first_dir_sector: 2,
            first_mini_fat: u32::MAX - 1,
            num_mini_fat: 0,
            first_difat: u32::MAX - 1,
            num_difat: 0,
            difat_entry_count: 1,
            sector_count: 4,
            fat_entry_count: 128,
            fat_free_count: 124,
            occupied_fat_entries: 4,
            fat_end_of_chain_count: 3,
            fat_reserved_count: 1,
            mini_fat_entry_count: 0,
            role_counts: vec![
                RoleCount {
                    role: "end-of-chain".to_string(),
                    count: 1,
                },
                RoleCount {
                    role: "special:directory".to_string(),
                    count: 1,
                },
            ],
            shared_sector_claims: vec![SharedSectorClaim {
                sector: 0,
                owners: vec!["WordDocument".to_string(), "VBA/PROJECT".to_string()],
            }],
            shared_chain_overlaps: vec![SharedChainOverlap {
                owners: vec!["WordDocument".to_string(), "VBA/PROJECT".to_string()],
                sectors: vec![0, 1],
                severity: "high".to_string(),
            }],
            start_sector_reuse: vec![StartSectorReuse {
                sector: 0,
                owners: vec!["WordDocument".to_string(), "VBA/PROJECT".to_string()],
            }],
            truncated_chain_counts: vec![TruncatedChainCount {
                bucket: "fat:WordDocument".to_string(),
                count: 1,
            }],
            structural_incoherence_counts: vec![StructuralIncoherenceCount {
                bucket: "directory:marked-free".to_string(),
                severity: "high".to_string(),
                count: 1,
            }],
            chain_health_by_root: vec![ChainHealthCount {
                bucket: "health:shared:root:WordDocument".to_string(),
                count: 1,
            }],
            chain_health_by_allocation: vec![ChainHealthCount {
                bucket: "health:shared:allocation:mini-fat".to_string(),
                count: 1,
            }],
            sector_overview: vec![SectorOverviewEntry {
                sector: 2,
                fat_raw: u32::MAX - 1,
                fat_value: "end-of-chain".to_string(),
                role: "end-of-chain".to_string(),
                special_roles: vec!["directory".to_string()],
                owners: Vec::new(),
            }],
            streams: vec![StreamSectorMap {
                path: "WordDocument".to_string(),
                logical_root: "WordDocument".to_string(),
                allocation: "mini-fat".to_string(),
                chain_state: "complete".to_string(),
                stream_health: "shared".to_string(),
                stream_risk: "high".to_string(),
                chain_terminal_raw: u32::MAX - 1,
                chain_terminal: "end-of-chain".to_string(),
                size_bytes: 64,
                start_sector: 0,
                expected_chain_len: 1,
                sector_chain: vec![0],
                chain_steps: vec![ChainStep {
                    sector: 0,
                    next_raw: u32::MAX - 1,
                    next: "end-of-chain".to_string(),
                }],
                sector_count: 1,
            }],
            anomalies: vec![SectorAnomaly {
                kind: "shared-sector".to_string(),
                severity: "high".to_string(),
                message: "example anomaly".to_string(),
            }],
        };
        let text = format_inspection_text(&inspection);
        assert!(text.contains("Header FAT Sectors: 1"));
        assert!(text.contains("Sector Score: high"));
        assert!(text.contains("DIFAT Entries: 1"));
        assert!(text.contains("FAT Reserved: 1"));
        assert!(text.contains("FAT Occupied: 4"));
        assert!(text.contains("Role Counts:"));
        assert!(text.contains("special:directory: 1"));
        assert!(text.contains("Shared Sectors:"));
        assert!(text.contains("0: WordDocument, VBA/PROJECT"));
        assert!(text.contains("Shared Chains:"));
        assert!(text.contains("WordDocument, VBA/PROJECT [high]: [0, 1]"));
        assert!(text.contains("Start Sector Reuse:"));
        assert!(text.contains("0: WordDocument, VBA/PROJECT"));
        assert!(text.contains("Truncated Chains:"));
        assert!(text.contains("fat:WordDocument: 1"));
        assert!(text.contains("Chain Health By Root:"));
        assert!(text.contains("health:shared:root:WordDocument: 1"));
        assert!(text.contains("Chain Health By Allocation:"));
        assert!(text.contains("health:shared:allocation:mini-fat: 1"));
        assert!(text.contains("Structural Incoherence:"));
        assert!(text.contains("directory:marked-free: 1 [high]"));
        assert!(text.contains("Sector Overview:"));
        assert!(text.contains("2: end-of-chain"));
        assert!(text.contains("Special Roles: directory"));
        assert!(text.contains("Logical Root: WordDocument"));
        assert!(text.contains("Chain State: complete"));
        assert!(text.contains("Stream Health: shared"));
        assert!(text.contains("Stream Risk: high"));
        assert!(text.contains("Size: 64 bytes"));
        assert!(text.contains("Start Sector: 0"));
        assert!(text.contains("Expected Chain: 1"));
        assert!(text.contains("Chain Terminal: end-of-chain"));
        assert!(text.contains("Chain: [0]"));
        assert!(text.contains("Steps:"));
        assert!(text.contains("0: 4294967294 (end-of-chain)"));
        assert!(text.contains("Anomalies:"));
        assert!(text.contains("shared-sector [high]: example anomaly"));
    }
}
