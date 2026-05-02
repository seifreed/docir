use crate::io_support::read_bounded_file;
use crate::severity::max_severity;
use crate::{AppResult, ParserConfig};
use docir_parser::ole::Cfb;
use serde::Serialize;
use std::path::Path;

/// Low-level summary of CFB sector allocation.
#[derive(Debug, Clone, Serialize)]
pub struct SectorInspection {
    pub container: String,
    pub sector_score: String,
    pub sector_size: u32,
    pub mini_sector_size: u32,
    pub mini_cutoff: u32,
    pub num_fat_sectors: u32,
    pub first_dir_sector: u32,
    pub first_mini_fat: u32,
    pub num_mini_fat: u32,
    pub first_difat: u32,
    pub num_difat: u32,
    pub difat_entry_count: usize,
    pub sector_count: u32,
    pub fat_entry_count: usize,
    pub fat_free_count: usize,
    pub occupied_fat_entries: usize,
    pub fat_end_of_chain_count: usize,
    pub fat_reserved_count: usize,
    pub mini_fat_entry_count: usize,
    pub role_counts: Vec<RoleCount>,
    pub shared_sector_claims: Vec<SharedSectorClaim>,
    pub shared_chain_overlaps: Vec<SharedChainOverlap>,
    pub start_sector_reuse: Vec<StartSectorReuse>,
    pub truncated_chain_counts: Vec<TruncatedChainCount>,
    pub structural_incoherence_counts: Vec<StructuralIncoherenceCount>,
    pub chain_health_by_root: Vec<ChainHealthCount>,
    pub chain_health_by_allocation: Vec<ChainHealthCount>,
    pub sector_overview: Vec<SectorOverviewEntry>,
    pub streams: Vec<StreamSectorMap>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub anomalies: Vec<SectorAnomaly>,
}

/// Sector-chain mapping for one stream.
#[derive(Debug, Clone, Serialize)]
pub struct StreamSectorMap {
    pub path: String,
    pub logical_root: String,
    pub allocation: String,
    pub chain_state: String,
    pub stream_health: String,
    pub stream_risk: String,
    pub chain_terminal_raw: u32,
    pub chain_terminal: String,
    pub size_bytes: u64,
    pub start_sector: u32,
    pub expected_chain_len: usize,
    pub sector_chain: Vec<u32>,
    pub chain_steps: Vec<ChainStep>,
    pub sector_count: u32,
}

#[derive(Debug, Clone, Serialize)]
pub struct ChainStep {
    pub sector: u32,
    pub next_raw: u32,
    pub next: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct SectorOverviewEntry {
    pub sector: u32,
    pub fat_raw: u32,
    pub fat_value: String,
    pub role: String,
    pub special_roles: Vec<String>,
    pub owners: Vec<SectorOwnerRef>,
}

#[derive(Debug, Clone, Serialize)]
pub struct SectorOwnerRef {
    pub path: String,
    pub index_in_chain: usize,
    pub is_terminal: bool,
}

#[derive(Debug, Clone, Serialize)]
pub struct RoleCount {
    pub role: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct SharedSectorClaim {
    pub sector: u32,
    pub owners: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct SharedChainOverlap {
    pub owners: Vec<String>,
    pub sectors: Vec<u32>,
    pub severity: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct TruncatedChainCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct StructuralIncoherenceCount {
    pub bucket: String,
    pub severity: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct StartSectorReuse {
    pub sector: u32,
    pub owners: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct ChainHealthCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct SectorAnomaly {
    pub kind: String,
    pub severity: String,
    pub message: String,
}

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

    let fat_free_count = cfb.fat_free_count();
    let fat_entry_count = cfb.fat_entry_count();
    let occupied_fat_entries = fat_entry_count.saturating_sub(fat_free_count);
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

fn collect_cfb_anomalies(
    cfb: &Cfb,
    streams: &[StreamSectorMap],
    fat_entry_count: usize,
    shared_sector_claims: &[SharedSectorClaim],
    shared_chain_overlaps: &[SharedChainOverlap],
    start_sector_reuse: &[StartSectorReuse],
    structural_incoherence_counts: &[StructuralIncoherenceCount],
) -> Vec<SectorAnomaly> {
    let mut anomalies = Vec::new();
    if cfb.num_fat_sectors() as usize > cfb.difat_entry_count() {
        anomalies.push(SectorAnomaly {
            kind: "difat-mismatch".to_string(),
            severity: "medium".to_string(),
            message: format!(
                "header declares {} FAT sector(s) but DIFAT exposes only {} entrie(s)",
                cfb.num_fat_sectors(),
                cfb.difat_entry_count()
            ),
        });
    }
    if fat_entry_count < cfb.sector_count() as usize {
        anomalies.push(SectorAnomaly {
            kind: "fat-shorter-than-sector-count".to_string(),
            severity: "high".to_string(),
            message: format!(
                "fat entry table shorter than sector count ({} < {})",
                fat_entry_count,
                cfb.sector_count()
            ),
        });
    }
    if cfb.mini_fat_entry_count() > 0
        && !streams.iter().any(|stream| stream.allocation == "mini-fat")
    {
        anomalies.push(SectorAnomaly {
            kind: "mini-fat-without-consumers".to_string(),
            severity: "medium".to_string(),
            message: "mini FAT is present but no stream currently resolves through mini-fat"
                .to_string(),
        });
    }
    for claim in shared_sector_claims {
        anomalies.push(SectorAnomaly {
            kind: "shared-sector".to_string(),
            severity: "high".to_string(),
            message: format!(
                "sector {} claimed by multiple streams: {}",
                claim.sector,
                claim.owners.join(", ")
            ),
        });
    }
    for overlap in shared_chain_overlaps {
        anomalies.push(SectorAnomaly {
            kind: "shared-chain".to_string(),
            severity: "high".to_string(),
            message: format!(
                "streams {} share chain segment {:?}",
                overlap.owners.join(", "),
                overlap.sectors
            ),
        });
    }
    for reuse in start_sector_reuse {
        anomalies.push(SectorAnomaly {
            kind: "start-sector-reuse".to_string(),
            severity: "medium".to_string(),
            message: format!(
                "start sector {} reused by streams: {}",
                reuse.sector,
                reuse.owners.join(", ")
            ),
        });
    }
    for entry in structural_incoherence_counts {
        anomalies.push(SectorAnomaly {
            kind: entry.bucket.clone(),
            severity: entry.severity.clone(),
            message: format!("{}: {}", entry.bucket, entry.count),
        });
    }
    anomalies
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

fn build_role_counts(sector_overview: &[SectorOverviewEntry]) -> Vec<RoleCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for entry in sector_overview {
        *counts.entry(entry.role.clone()).or_insert(0) += 1;
        for special in &entry.special_roles {
            *counts.entry(format!("special:{special}")).or_insert(0) += 1;
        }
    }
    counts
        .into_iter()
        .map(|(role, count)| RoleCount { role, count })
        .collect()
}

fn build_shared_sector_claims(sector_overview: &[SectorOverviewEntry]) -> Vec<SharedSectorClaim> {
    sector_overview
        .iter()
        .filter(|entry| entry.owners.len() > 1)
        .map(|entry| SharedSectorClaim {
            sector: entry.sector,
            owners: entry
                .owners
                .iter()
                .map(|owner| owner.path.clone())
                .collect(),
        })
        .collect()
}

fn build_shared_chain_overlaps(sector_overview: &[SectorOverviewEntry]) -> Vec<SharedChainOverlap> {
    let mut grouped = std::collections::BTreeMap::<Vec<String>, Vec<u32>>::new();
    for entry in sector_overview
        .iter()
        .filter(|entry| entry.owners.len() > 1)
    {
        let owners = entry
            .owners
            .iter()
            .map(|owner| owner.path.clone())
            .collect::<Vec<_>>();
        grouped.entry(owners).or_default().push(entry.sector);
    }

    let mut overlaps = Vec::new();
    for (owners, mut sectors) in grouped {
        sectors.sort_unstable();
        let mut run = Vec::new();
        for sector in sectors {
            if run.last().is_none_or(|last| *last + 1 == sector) {
                run.push(sector);
            } else {
                if run.len() >= 2 {
                    overlaps.push(SharedChainOverlap {
                        owners: owners.clone(),
                        sectors: run.clone(),
                        severity: shared_chain_severity(&owners, &run).to_string(),
                    });
                }
                run.clear();
                run.push(sector);
            }
        }
        if run.len() >= 2 {
            overlaps.push(SharedChainOverlap {
                severity: shared_chain_severity(&owners, &run).to_string(),
                owners,
                sectors: run,
            });
        }
    }
    overlaps
}

fn shared_chain_severity(owners: &[String], sectors: &[u32]) -> &'static str {
    let touches_sensitive = owners.iter().any(|owner| {
        let upper = owner.to_ascii_uppercase();
        upper == "WORDDOCUMENT"
            || upper == "WORKBOOK"
            || upper == "POWERPOINT DOCUMENT"
            || upper.starts_with("VBA/")
    });
    if sectors.len() >= 3 || touches_sensitive {
        "high"
    } else {
        "medium"
    }
}

fn build_start_sector_reuse(streams: &[StreamSectorMap]) -> Vec<StartSectorReuse> {
    let mut grouped = std::collections::BTreeMap::<u32, Vec<String>>::new();
    for stream in streams {
        if stream.start_sector != u32::MAX {
            grouped
                .entry(stream.start_sector)
                .or_default()
                .push(stream.path.clone());
        }
    }
    grouped
        .into_iter()
        .filter(|(_, owners)| owners.len() > 1)
        .map(|(sector, owners)| StartSectorReuse { sector, owners })
        .collect()
}

fn build_truncated_chain_counts(streams: &[StreamSectorMap]) -> Vec<TruncatedChainCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for stream in streams {
        if stream.chain_state == "truncated" {
            *counts
                .entry(format!("{}:{}", stream.allocation, stream.logical_root))
                .or_insert(0) += 1;
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| TruncatedChainCount { bucket, count })
        .collect()
}

fn build_structural_incoherence_counts(
    sector_overview: &[SectorOverviewEntry],
    streams: &[StreamSectorMap],
    mini_fat_entry_count: usize,
) -> Vec<StructuralIncoherenceCount> {
    let mut counts = std::collections::BTreeMap::<String, (String, usize)>::new();
    if mini_fat_entry_count > 0 && !streams.iter().any(|stream| stream.allocation == "mini-fat") {
        counts.insert(
            "mini-fat-without-consumers".to_string(),
            ("medium".to_string(), 1),
        );
    }
    for entry in sector_overview {
        for special in &entry.special_roles {
            if entry.role == "free" {
                let slot = counts
                    .entry(format!("{special}:marked-free"))
                    .or_insert(("high".to_string(), 0));
                slot.1 += 1;
            }
            if !entry.owners.is_empty() {
                let slot = counts
                    .entry(format!("{special}:claimed-by-stream"))
                    .or_insert(("medium".to_string(), 0));
                slot.1 += 1;
            }
        }
    }
    for stream in streams {
        let has_mismatched_start = stream.size_bytes > 0
            && ((!stream.sector_chain.is_empty() && stream.start_sector != stream.sector_chain[0])
                || (stream.start_sector != u32::MAX
                    && (stream.sector_chain.is_empty() || stream.chain_terminal_raw == u32::MAX)));
        if has_mismatched_start {
            let slot = counts
                .entry(format!("start-sector-mismatch:{}", stream.path))
                .or_insert(("high".to_string(), 0));
            slot.1 += 1;
        }
    }
    counts
        .into_iter()
        .map(|(bucket, (severity, count))| StructuralIncoherenceCount {
            bucket,
            severity,
            count,
        })
        .collect()
}

fn stream_health_tags(
    streams: &[StreamSectorMap],
    shared_sector_claims: &[SharedSectorClaim],
    start_sector_reuse: &[StartSectorReuse],
) -> Vec<(String, String, Vec<String>)> {
    let mut out = Vec::new();
    for stream in streams {
        let mut tags = vec![stream.chain_state.clone()];
        let is_shared = shared_sector_claims
            .iter()
            .any(|claim| claim.owners.iter().any(|owner| owner == &stream.path));
        if is_shared {
            tags.push("shared".to_string());
        }
        if start_sector_reuse
            .iter()
            .any(|reuse| reuse.owners.iter().any(|owner| owner == &stream.path))
        {
            tags.push("start-reused".to_string());
        }
        let is_invalid_start = stream.size_bytes > 0
            && stream.start_sector != u32::MAX
            && (stream.start_sector >= stream.sector_count
                || (!stream.sector_chain.is_empty()
                    && stream.start_sector != stream.sector_chain[0])
                || (stream.sector_chain.is_empty() && stream.chain_terminal_raw != u32::MAX - 1));
        if is_invalid_start {
            tags.push("invalid-start".to_string());
        }
        out.push((stream.logical_root.clone(), stream.allocation.clone(), tags));
    }
    out
}

fn build_chain_health_counts_by_root(
    streams: &[StreamSectorMap],
    shared_sector_claims: &[SharedSectorClaim],
    start_sector_reuse: &[StartSectorReuse],
) -> Vec<ChainHealthCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for (logical_root, _, tags) in
        stream_health_tags(streams, shared_sector_claims, start_sector_reuse)
    {
        for tag in tags {
            *counts
                .entry(format!("health:{tag}:root:{logical_root}"))
                .or_insert(0) += 1;
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| ChainHealthCount { bucket, count })
        .collect()
}

fn build_chain_health_counts_by_allocation(
    streams: &[StreamSectorMap],
    shared_sector_claims: &[SharedSectorClaim],
    start_sector_reuse: &[StartSectorReuse],
) -> Vec<ChainHealthCount> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for (_, allocation, tags) in
        stream_health_tags(streams, shared_sector_claims, start_sector_reuse)
    {
        for tag in tags {
            *counts
                .entry(format!("health:{tag}:allocation:{allocation}"))
                .or_insert(0) += 1;
        }
    }
    counts
        .into_iter()
        .map(|(bucket, count)| ChainHealthCount { bucket, count })
        .collect()
}

fn build_sector_score(streams: &[StreamSectorMap], anomalies: &[SectorAnomaly]) -> String {
    if anomalies.iter().any(|anomaly| anomaly.severity == "high")
        || streams.iter().any(|stream| stream.stream_risk == "high")
    {
        "high".to_string()
    } else if anomalies.iter().any(|anomaly| anomaly.severity == "medium")
        || streams.iter().any(|stream| stream.stream_risk == "medium")
    {
        "medium".to_string()
    } else if !anomalies.is_empty() || streams.iter().any(|stream| stream.stream_risk == "low") {
        "low".to_string()
    } else {
        "none".to_string()
    }
}

fn apply_stream_health(
    streams: &mut [StreamSectorMap],
    shared_sector_claims: &[SharedSectorClaim],
    shared_chain_overlaps: &[SharedChainOverlap],
    start_sector_reuse: &[StartSectorReuse],
) {
    for stream in streams {
        let mut severities = vec!["none".to_string()];
        let mut health = stream.chain_state.as_str();
        let mut is_shared = false;
        if shared_sector_claims
            .iter()
            .any(|claim| claim.owners.iter().any(|owner| owner == &stream.path))
        {
            severities.push("high".to_string());
            is_shared = true;
        }
        if let Some(overlap) = shared_chain_overlaps
            .iter()
            .find(|overlap| overlap.owners.iter().any(|owner| owner == &stream.path))
        {
            severities.push(overlap.severity.clone());
            is_shared = true;
        }
        if start_sector_reuse
            .iter()
            .any(|reuse| reuse.owners.iter().any(|owner| owner == &stream.path))
        {
            severities.push("medium".to_string());
            if health == "complete" {
                health = "start-reused";
            }
        }
        let is_invalid_start = stream.size_bytes > 0
            && stream.start_sector != u32::MAX
            && (stream.start_sector >= stream.sector_count
                || (!stream.sector_chain.is_empty()
                    && stream.start_sector != stream.sector_chain[0])
                || (stream.sector_chain.is_empty() && stream.chain_terminal_raw != u32::MAX - 1));
        if is_invalid_start {
            severities.push("high".to_string());
            health = "invalid-start";
        } else if is_shared {
            health = "shared";
        }
        if stream.chain_state == "truncated" {
            severities.push("medium".to_string());
            if health == "complete" {
                health = "truncated";
            }
        }
        stream.stream_health = health.to_string();
        stream.stream_risk = max_severity(severities);
    }
}

#[cfg(test)]
mod tests;
