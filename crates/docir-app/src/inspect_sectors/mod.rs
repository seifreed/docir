use crate::io_support::read_bounded_file;
use crate::{AppResult, ParserConfig};
use docir_parser::ole::Cfb;
use std::path::Path;

mod anomalies;
mod counts;
mod health;
mod types;

pub use types::*;

use anomalies::collect_cfb_anomalies;
use counts::{
    build_role_counts, build_shared_chain_overlaps, build_shared_sector_claims,
    build_start_sector_reuse, build_structural_incoherence_counts, build_truncated_chain_counts,
};
use health::{
    apply_stream_health, build_chain_health_counts_by_allocation,
    build_chain_health_counts_by_root, build_sector_score,
};

/// Inspect sector allocation from a legacy CFB/OLE file on disk.
pub fn inspect_sectors_path<P: AsRef<Path>>(
    path: P,
    config: &ParserConfig,
) -> AppResult<SectorInspection> {
    let bytes = read_bounded_file(path, config.max_input_size)?;
    inspect_sectors_bytes(&bytes)
}

/// Inspect sector allocation from raw CFB/OLE bytes.
pub fn inspect_sectors_bytes(data: &[u8]) -> AppResult<SectorInspection> {
    let cfb = Cfb::parse(data.to_vec())?;
    let mut anomalies = Vec::new();
    let mut streams = build_stream_maps(&cfb, &mut anomalies);
    streams.sort_by(|left, right| left.path.cmp(&right.path));
    let (sector_owners, referenced_sectors) = build_sector_owner_map(&streams);
    let (directory_sectors, mini_fat_sectors, difat_sectors) = collect_system_chains(&cfb);
    let sector_overview = build_sector_overview_entries(
        &cfb,
        &referenced_sectors,
        &directory_sectors,
        &mini_fat_sectors,
        &difat_sectors,
        sector_owners,
    );
    let role_counts = build_role_counts(&sector_overview);
    let shared_sector_claims = build_shared_sector_claims(&sector_overview);
    let shared_chain_overlaps = build_shared_chain_overlaps(&sector_overview);
    let start_sector_reuse = build_start_sector_reuse(&streams);
    let truncated_chain_counts = build_truncated_chain_counts(&streams);
    apply_stream_health(
        &mut streams,
        &shared_sector_claims,
        &shared_chain_overlaps,
        &start_sector_reuse,
    );
    let fat_entry_count = cfb.fat_entry_count();
    let structural_incoherence_counts =
        build_structural_incoherence_counts(&sector_overview, &streams, cfb.mini_fat_entry_count());
    anomalies.extend(collect_cfb_anomalies(
        &cfb,
        &streams,
        fat_entry_count,
        &shared_sector_claims,
        &shared_chain_overlaps,
        &start_sector_reuse,
        &structural_incoherence_counts,
    ));
    let chain_health_by_root =
        build_chain_health_counts_by_root(&streams, &shared_sector_claims, &start_sector_reuse);
    let chain_health_by_allocation = build_chain_health_counts_by_allocation(
        &streams,
        &shared_sector_claims,
        &start_sector_reuse,
    );
    let sector_score = build_sector_score(&streams, &anomalies);
    let fat_free_count = cfb.fat_free_count();
    let occupied_fat_entries = fat_entry_count.saturating_sub(fat_free_count);
    Ok(SectorInspection {
        container: "cfb-ole".to_string(),
        sector_score,
        sector_size: cfb.sector_size(),
        mini_sector_size: cfb.mini_sector_size(),
        mini_cutoff: cfb.mini_cutoff(),
        num_fat_sectors: cfb.num_fat_sectors(),
        first_dir_sector: cfb.first_dir_sector(),
        first_mini_fat: cfb.first_mini_fat(),
        num_mini_fat: cfb.num_mini_fat(),
        first_difat: cfb.first_difat(),
        num_difat: cfb.num_difat(),
        difat_entry_count: cfb.difat_entry_count(),
        sector_count: cfb.sector_count(),
        fat_entry_count,
        fat_free_count,
        occupied_fat_entries,
        fat_end_of_chain_count: cfb.fat_end_of_chain_count(),
        fat_reserved_count: cfb.fat_reserved_count(),
        mini_fat_entry_count: cfb.mini_fat_entry_count(),
        role_counts,
        shared_sector_claims,
        shared_chain_overlaps,
        start_sector_reuse,
        truncated_chain_counts,
        structural_incoherence_counts,
        chain_health_by_root,
        chain_health_by_allocation,
        sector_overview,
        streams,
        anomalies,
    })
}

fn build_stream_maps(cfb: &Cfb, anomalies: &mut Vec<SectorAnomaly>) -> Vec<StreamSectorMap> {
    cfb.list_streams()
        .into_iter()
        .map(|path| {
            let uses_mini_fat = cfb.stream_uses_mini_fat(&path).unwrap_or(false);
            let allocation = if uses_mini_fat { "mini-fat" } else { "fat" }.to_string();
            let size_bytes = cfb.stream_size(&path).unwrap_or(0);
            let start_sector = cfb
                .entry_metadata(&path)
                .map(|entry| entry.start_sector)
                .unwrap_or(u32::MAX);
            let chain_terminal_raw = cfb.stream_chain_terminal(&path).unwrap_or(u32::MAX);
            let unit_size = if uses_mini_fat {
                cfb.mini_sector_size() as u64
            } else {
                cfb.sector_size() as u64
            };
            let expected_chain_len = if size_bytes == 0 {
                0
            } else {
                size_bytes.div_ceil(unit_size) as usize
            };
            let sector_chain = cfb.stream_sector_chain(&path).unwrap_or_default();
            let chain_state = classify_chain_state(size_bytes, expected_chain_len, &sector_chain);
            if size_bytes > 0 && sector_chain.is_empty() {
                anomalies.push(SectorAnomaly {
                    kind: "missing-chain".to_string(),
                    severity: "high".to_string(),
                    message: format!(
                        "stream {path} has non-zero size ({size_bytes}) but no sector chain"
                    ),
                });
            } else if expected_chain_len != sector_chain.len() {
                anomalies.push(SectorAnomaly {
                    kind: "truncated-chain".to_string(),
                    severity: "medium".to_string(),
                    message: format!(
                        "stream {path} expected {expected_chain_len} {allocation} sector(s) but resolved {}",
                        sector_chain.len()
                    ),
                });
            }
            StreamSectorMap {
                logical_root: logical_root(&path),
                allocation,
                chain_state,
                stream_health: "complete".to_string(),
                stream_risk: "none".to_string(),
                chain_terminal_raw,
                chain_terminal: describe_chain_terminal(chain_terminal_raw).to_string(),
                size_bytes,
                start_sector,
                expected_chain_len,
                chain_steps: build_chain_steps(&sector_chain, chain_terminal_raw),
                sector_chain,
                path,
                sector_count: cfb.sector_count(),
            }
        })
        .collect::<Vec<_>>()
}

fn build_sector_owner_map(
    streams: &[StreamSectorMap],
) -> (
    std::collections::HashMap<u32, Vec<SectorOwnerRef>>,
    std::collections::HashSet<u32>,
) {
    let mut sector_owners = std::collections::HashMap::<u32, Vec<SectorOwnerRef>>::new();
    for stream in streams {
        for (index_in_chain, sector) in stream.sector_chain.iter().enumerate() {
            sector_owners
                .entry(*sector)
                .or_default()
                .push(SectorOwnerRef {
                    path: stream.path.clone(),
                    index_in_chain,
                    is_terminal: index_in_chain + 1 == stream.sector_chain.len(),
                });
        }
    }
    let referenced_sectors = sector_owners
        .keys()
        .copied()
        .collect::<std::collections::HashSet<_>>();
    (sector_owners, referenced_sectors)
}

fn collect_system_chains(
    cfb: &Cfb,
) -> (
    std::collections::HashSet<u32>,
    std::collections::HashSet<u32>,
    std::collections::HashSet<u32>,
) {
    let directory_sectors = collect_fat_chain(cfb, cfb.first_dir_sector(), cfb.sector_count());
    let mini_fat_sectors = if cfb.num_mini_fat() > 0 {
        collect_fat_chain(cfb, cfb.first_mini_fat(), cfb.sector_count())
    } else {
        std::collections::HashSet::new()
    };
    let difat_sectors = if cfb.num_difat() > 0 {
        collect_fat_chain(cfb, cfb.first_difat(), cfb.sector_count())
    } else {
        std::collections::HashSet::new()
    };
    (directory_sectors, mini_fat_sectors, difat_sectors)
}

fn build_sector_overview_entries(
    cfb: &Cfb,
    referenced_sectors: &std::collections::HashSet<u32>,
    directory_sectors: &std::collections::HashSet<u32>,
    mini_fat_sectors: &std::collections::HashSet<u32>,
    difat_sectors: &std::collections::HashSet<u32>,
    mut sector_owners: std::collections::HashMap<u32, Vec<SectorOwnerRef>>,
) -> Vec<SectorOverviewEntry> {
    (0..cfb.sector_count())
        .map(|sector| {
            let fat_raw = cfb.fat_entry_value(sector).unwrap_or(u32::MAX);
            SectorOverviewEntry {
                sector,
                fat_raw,
                fat_value: describe_chain_terminal(fat_raw).to_string(),
                role: classify_sector_role(sector, fat_raw, referenced_sectors).to_string(),
                special_roles: collect_special_roles(
                    sector,
                    fat_raw,
                    directory_sectors,
                    mini_fat_sectors,
                    difat_sectors,
                ),
                owners: sector_owners.remove(&sector).unwrap_or_default(),
            }
        })
        .collect::<Vec<_>>()
}

fn logical_root(path: &str) -> String {
    path.split('/').next().unwrap_or(path).to_string()
}

fn classify_chain_state(
    size_bytes: u64,
    expected_chain_len: usize,
    sector_chain: &[u32],
) -> String {
    if size_bytes == 0 {
        "empty".to_string()
    } else if sector_chain.len() == expected_chain_len && !sector_chain.is_empty() {
        "complete".to_string()
    } else {
        "truncated".to_string()
    }
}

fn describe_chain_terminal(value: u32) -> &'static str {
    match value {
        u32::MAX => "free-sect",
        val if val == u32::MAX - 1 => "end-of-chain",
        _ => "sector",
    }
}

fn build_chain_steps(sector_chain: &[u32], terminal_raw: u32) -> Vec<ChainStep> {
    sector_chain
        .iter()
        .enumerate()
        .map(|(idx, sector)| {
            let next_raw = sector_chain.get(idx + 1).copied().unwrap_or(terminal_raw);
            ChainStep {
                sector: *sector,
                next_raw,
                next: describe_chain_terminal(next_raw).to_string(),
            }
        })
        .collect()
}

fn classify_sector_role(
    sector: u32,
    fat_raw: u32,
    referenced_sectors: &std::collections::HashSet<u32>,
) -> &'static str {
    if referenced_sectors.contains(&sector) {
        "stream-data"
    } else if fat_raw == u32::MAX {
        "free"
    } else if fat_raw == u32::MAX - 1 {
        "end-of-chain"
    } else if fat_raw == u32::MAX - 3 {
        "fat"
    } else {
        "next-sector"
    }
}

fn collect_fat_chain(
    cfb: &Cfb,
    start_sector: u32,
    sector_count: u32,
) -> std::collections::HashSet<u32> {
    let mut out = std::collections::HashSet::new();
    let mut current = start_sector;
    let mut guard = 0u32;
    while current != u32::MAX && current != u32::MAX - 1 && guard < sector_count {
        if !out.insert(current) {
            break;
        }
        current = cfb.fat_entry_value(current).unwrap_or(u32::MAX - 1);
        guard += 1;
    }
    out
}

fn collect_special_roles(
    sector: u32,
    fat_raw: u32,
    directory_sectors: &std::collections::HashSet<u32>,
    mini_fat_sectors: &std::collections::HashSet<u32>,
    difat_sectors: &std::collections::HashSet<u32>,
) -> Vec<String> {
    let mut out = Vec::new();
    if directory_sectors.contains(&sector) {
        out.push("directory".to_string());
    }
    if mini_fat_sectors.contains(&sector) {
        out.push("mini-fat".to_string());
    }
    if difat_sectors.contains(&sector) {
        out.push("difat".to_string());
    }
    if fat_raw == u32::MAX - 3 {
        out.push("fat-table".to_string());
    }
    out
}

#[cfg(test)]
mod tests;
