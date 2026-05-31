use super::{format_inspection_text, run};
use crate::cli::JsonOutputOpts;
use crate::test_support;
use docir_app::{
    ChainHealthCount, ChainStep, ParserConfig, RoleCount, SectorAnomaly, SectorInspection,
    SectorOverviewEntry, SharedChainOverlap, SharedSectorClaim, StartSectorReuse, StreamSectorMap,
    StructuralIncoherenceCount, TruncatedChainCount, test_support::build_test_cfb,
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
    let text = format_inspection_text(&sector_inspection_fixture());

    assert_sector_header_text(&text);
    assert_sector_summary_text(&text);
    assert_sector_map_text(&text);
    assert_sector_anomaly_text(&text);
}

fn sector_inspection_fixture() -> SectorInspection {
    SectorInspection {
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
            role_count("end-of-chain", 1),
            role_count("special:directory", 1),
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
        truncated_chain_counts: vec![truncated_chain_count("fat:WordDocument", 1)],
        structural_incoherence_counts: vec![StructuralIncoherenceCount {
            bucket: "directory:marked-free".to_string(),
            severity: "high".to_string(),
            count: 1,
        }],
        chain_health_by_root: vec![chain_health_count("health:shared:root:WordDocument", 1)],
        chain_health_by_allocation: vec![chain_health_count(
            "health:shared:allocation:mini-fat",
            1,
        )],
        sector_overview: vec![sector_overview_fixture()],
        streams: vec![stream_sector_map_fixture()],
        anomalies: vec![SectorAnomaly {
            kind: "shared-sector".to_string(),
            severity: "high".to_string(),
            message: "example anomaly".to_string(),
        }],
    }
}

fn sector_overview_fixture() -> SectorOverviewEntry {
    SectorOverviewEntry {
        sector: 2,
        fat_raw: u32::MAX - 1,
        fat_value: "end-of-chain".to_string(),
        role: "end-of-chain".to_string(),
        special_roles: vec!["directory".to_string()],
        owners: Vec::new(),
    }
}

fn stream_sector_map_fixture() -> StreamSectorMap {
    StreamSectorMap {
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
    }
}

fn role_count(role: &str, count: usize) -> RoleCount {
    RoleCount {
        role: role.to_string(),
        count,
    }
}

fn truncated_chain_count(bucket: &str, count: usize) -> TruncatedChainCount {
    TruncatedChainCount {
        bucket: bucket.to_string(),
        count,
    }
}

fn chain_health_count(bucket: &str, count: usize) -> ChainHealthCount {
    ChainHealthCount {
        bucket: bucket.to_string(),
        count,
    }
}

fn assert_sector_header_text(text: &str) {
    assert_contains_all(
        text,
        &[
            "Header FAT Sectors: 1",
            "Sector Score: high",
            "DIFAT Entries: 1",
            "FAT Reserved: 1",
            "FAT Occupied: 4",
        ],
    );
}

fn assert_sector_summary_text(text: &str) {
    assert_contains_all(
        text,
        &[
            "Role Counts:",
            "special:directory: 1",
            "Shared Sectors:",
            "0: WordDocument, VBA/PROJECT",
            "Shared Chains:",
            "WordDocument, VBA/PROJECT [high]: [0, 1]",
            "Start Sector Reuse:",
            "Truncated Chains:",
            "fat:WordDocument: 1",
            "Chain Health By Root:",
            "health:shared:root:WordDocument: 1",
            "Chain Health By Allocation:",
            "health:shared:allocation:mini-fat: 1",
            "Structural Incoherence:",
            "directory:marked-free: 1 [high]",
        ],
    );
}

fn assert_sector_map_text(text: &str) {
    assert_contains_all(
        text,
        &[
            "Sector Overview:",
            "2: end-of-chain",
            "Special Roles: directory",
            "Logical Root: WordDocument",
            "Chain State: complete",
            "Stream Health: shared",
            "Stream Risk: high",
            "Size: 64 bytes",
            "Start Sector: 0",
            "Expected Chain: 1",
            "Chain Terminal: end-of-chain",
            "Chain: [0]",
            "Steps:",
            "0: 4294967294 (end-of-chain)",
        ],
    );
}

fn assert_sector_anomaly_text(text: &str) {
    assert_contains_all(
        text,
        &["Anomalies:", "shared-sector [high]: example anomaly"],
    );
}

fn assert_contains_all(text: &str, expected: &[&str]) {
    for fragment in expected {
        assert!(text.contains(fragment), "missing fragment: {fragment}");
    }
}
