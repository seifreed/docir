use crate::severity::max_severity;
use crate::{AppResult, ParserConfig};
use docir_parser::ole::Cfb;
use docir_parser::ParseError as ParserParseError;
use serde::Serialize;
use std::fs;
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
    inspect_sectors_bytes(&bytes)
}

/// Inspect sector allocation from raw CFB/OLE bytes.
pub fn inspect_sectors_bytes(data: &[u8]) -> AppResult<SectorInspection> {
    let cfb = Cfb::parse(data.to_vec())?;
    let mut anomalies = Vec::new();
    let mut streams = cfb
        .list_streams()
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
        .collect::<Vec<_>>();
    streams.sort_by(|left, right| left.path.cmp(&right.path));
    let mut sector_owners = std::collections::HashMap::<u32, Vec<SectorOwnerRef>>::new();
    for stream in &streams {
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
    let directory_sectors = collect_fat_chain(&cfb, cfb.first_dir_sector(), cfb.sector_count());
    let mini_fat_sectors = if cfb.num_mini_fat() > 0 {
        collect_fat_chain(&cfb, cfb.first_mini_fat(), cfb.sector_count())
    } else {
        std::collections::HashSet::new()
    };
    let difat_sectors = if cfb.num_difat() > 0 {
        collect_fat_chain(&cfb, cfb.first_difat(), cfb.sector_count())
    } else {
        std::collections::HashSet::new()
    };
    let sector_overview = (0..cfb.sector_count())
        .map(|sector| {
            let fat_raw = cfb.fat_entry_value(sector).unwrap_or(u32::MAX);
            SectorOverviewEntry {
                sector,
                fat_raw,
                fat_value: describe_chain_terminal(fat_raw).to_string(),
                role: classify_sector_role(sector, fat_raw, &referenced_sectors).to_string(),
                special_roles: collect_special_roles(
                    sector,
                    fat_raw,
                    &directory_sectors,
                    &mini_fat_sectors,
                    &difat_sectors,
                ),
                owners: sector_owners.remove(&sector).unwrap_or_default(),
            }
        })
        .collect::<Vec<_>>();
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
    for claim in &shared_sector_claims {
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
    for overlap in &shared_chain_overlaps {
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
    for reuse in &start_sector_reuse {
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
    let structural_incoherence_counts =
        build_structural_incoherence_counts(&sector_overview, &streams, cfb.mini_fat_entry_count());
    for entry in &structural_incoherence_counts {
        anomalies.push(SectorAnomaly {
            kind: entry.bucket.clone(),
            severity: entry.severity.clone(),
            message: format!("{}: {}", entry.bucket, entry.count),
        });
    }
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
mod tests {
    use super::inspect_sectors_bytes;
    use crate::inspect_directory_bytes;
    use crate::test_support::{
        build_test_cfb, patch_test_cfb_directory_entry, patch_test_cfb_fat_entry,
        patch_test_cfb_header_u32, TestCfbDirectoryPatch,
    };

    #[test]
    fn inspect_sectors_reads_fat_and_stream_chains() {
        let inspection = inspect_sectors_bytes(&build_test_cfb(&[
            ("WordDocument", b"doc"),
            ("VBA/PROJECT", b"meta"),
        ]))
        .expect("inspection");

        assert_eq!(inspection.container, "cfb-ole");
        assert_eq!(inspection.sector_size, 512);
        assert_eq!(inspection.sector_count, 4);
        assert_eq!(inspection.num_fat_sectors, 1);
        assert_eq!(inspection.first_dir_sector, 2);
        assert_eq!(inspection.num_difat, 0);
        assert_eq!(inspection.difat_entry_count, 109);
        assert_eq!(inspection.fat_reserved_count, 1);
        assert_eq!(inspection.occupied_fat_entries, 128);
        assert!(inspection.anomalies.is_empty());
        assert!(inspection
            .role_counts
            .iter()
            .any(|entry| entry.role == "stream-data" && entry.count >= 1));
        assert!(inspection
            .role_counts
            .iter()
            .any(|entry| entry.role == "special:directory" && entry.count >= 1));
        assert!(inspection
            .sector_overview
            .iter()
            .any(|entry| entry.sector == 0
                && entry.role == "stream-data"
                && entry.fat_value == "end-of-chain"
                && entry.special_roles.is_empty()
                && entry.owners.len() == 1
                && entry.owners[0].path == "WordDocument"
                && entry.owners[0].index_in_chain == 0
                && entry.owners[0].is_terminal));
        assert!(inspection.sector_overview.iter().any(
            |entry| entry.sector == 2 && entry.special_roles.contains(&"directory".to_string())
        ));
        assert!(inspection
            .streams
            .iter()
            .any(|stream| stream.path == "WordDocument"
                && stream.sector_chain == vec![0]
                && stream.chain_state == "complete"
                && stream.chain_terminal_raw == u32::MAX - 1
                && stream.chain_terminal == "end-of-chain"
                && stream.start_sector == 0
                && stream.logical_root == "WordDocument"
                && stream.size_bytes == 3
                && stream.expected_chain_len == 1
                && stream.chain_steps.len() == 1
                && stream.chain_steps[0].sector == 0
                && stream.chain_steps[0].next_raw == u32::MAX - 1));
        assert!(inspection
            .streams
            .iter()
            .any(|stream| stream.path == "VBA/PROJECT" && stream.sector_chain == vec![1]));
    }

    #[test]
    fn inspect_sectors_detects_shared_sector_claims() {
        let base = build_test_cfb(&[("WordDocument", b"doc"), ("VBA/PROJECT", b"meta")]);
        let inspection = inspect_directory_bytes(&base).expect("directory");
        let vba_entry = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "VBA/PROJECT")
            .expect("vba entry");
        let patched = patch_test_cfb_directory_entry(
            &base,
            vba_entry.entry_index,
            TestCfbDirectoryPatch {
                start_sector: Some(0),
                ..Default::default()
            },
        );

        let inspection = inspect_sectors_bytes(&patched).expect("inspection");
        assert!(inspection
            .shared_sector_claims
            .iter()
            .any(|claim| claim.sector == 0 && claim.owners.len() >= 2));
        assert!(inspection
            .anomalies
            .iter()
            .any(|value| value.kind == "shared-sector"
                && value.message.contains("claimed by multiple streams")
                && value.severity == "high"));
    }

    #[test]
    fn inspect_sectors_detects_shared_chain_overlaps() {
        let base = build_test_cfb(&[
            ("WordDocument", &[0u8; 1025]),
            ("VBA/PROJECT", &[1u8; 1025]),
        ]);
        let inspection = inspect_directory_bytes(&base).expect("directory");
        let vba_entry = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "VBA/PROJECT")
            .expect("vba entry");
        let patched = patch_test_cfb_directory_entry(
            &base,
            vba_entry.entry_index,
            TestCfbDirectoryPatch {
                start_sector: Some(0),
                ..Default::default()
            },
        );

        let inspection = inspect_sectors_bytes(&patched).expect("inspection");
        assert!(inspection
            .shared_chain_overlaps
            .iter()
            .any(|overlap| overlap.sectors.len() >= 2));
    }

    #[test]
    fn inspect_sectors_summarizes_truncated_chains_by_logical_root() {
        let base = build_test_cfb(&[("WordDocument", &[0u8; 600]), ("VBA/PROJECT", b"meta")]);
        let patched = patch_test_cfb_fat_entry(&base, 0, u32::MAX - 1);

        let inspection = inspect_sectors_bytes(&patched).expect("inspection");
        assert!(inspection
            .truncated_chain_counts
            .iter()
            .any(|entry| entry.bucket == "fat:WordDocument" && entry.count == 1));
        assert!(inspection
            .streams
            .iter()
            .any(|stream| stream.path == "WordDocument" && stream.chain_state == "truncated"));
    }

    #[test]
    fn inspect_sectors_summarizes_structural_incoherence() {
        let base = build_test_cfb(&[("WordDocument", b"doc")]);
        let patched =
            patch_test_cfb_header_u32(&patch_test_cfb_header_u32(&base, 0x3C, 0), 0x40, 1);

        let inspection = inspect_sectors_bytes(&patched).expect("inspection");
        assert!(inspection
            .structural_incoherence_counts
            .iter()
            .any(|entry| entry.bucket == "mini-fat-without-consumers"
                && entry.severity == "medium"
                && entry.count == 1));
    }

    #[test]
    fn inspect_sectors_detects_start_sector_mismatch() {
        let base = build_test_cfb(&[("WordDocument", &[0u8; 600])]);
        let inspection = inspect_directory_bytes(&base).expect("directory");
        let word_entry = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "WordDocument")
            .expect("word entry");
        let patched = patch_test_cfb_fat_entry(&base, word_entry.start_sector, 0xFFFF_FFFF);

        let inspection = inspect_sectors_bytes(&patched).expect("inspection");
        assert!(inspection
            .structural_incoherence_counts
            .iter()
            .any(
                |entry| entry.bucket.contains("start-sector-mismatch:WordDocument")
                    && entry.severity == "high"
            ));
    }

    #[test]
    fn inspect_sectors_detects_start_sector_reuse_and_chain_health() {
        let base = build_test_cfb(&[("WordDocument", b"doc"), ("VBA/PROJECT", b"meta")]);
        let inspection = inspect_directory_bytes(&base).expect("directory");
        let vba_entry = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "VBA/PROJECT")
            .expect("vba entry");
        let patched = patch_test_cfb_directory_entry(
            &base,
            vba_entry.entry_index,
            TestCfbDirectoryPatch {
                start_sector: Some(0),
                ..Default::default()
            },
        );

        let inspection = inspect_sectors_bytes(&patched).expect("inspection");
        assert!(inspection
            .start_sector_reuse
            .iter()
            .any(|entry| entry.sector == 0 && entry.owners.len() >= 2));
        assert!(inspection
            .chain_health_by_root
            .iter()
            .any(|entry| entry.bucket.contains("health:start-reused:root:") && entry.count >= 1));
    }

    #[test]
    fn inspect_sectors_assigns_high_risk_to_shared_and_invalid_start_streams() {
        let base = build_test_cfb(&[
            ("WordDocument", &[0u8; 1025]),
            ("VBA/PROJECT", &[1u8; 1025]),
        ]);
        let inspection = inspect_directory_bytes(&base).expect("directory");
        let word_entry = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "WordDocument")
            .expect("word entry");
        let vba_entry = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "VBA/PROJECT")
            .expect("vba entry");
        let patched = patch_test_cfb_directory_entry(
            &patch_test_cfb_directory_entry(
                &base,
                vba_entry.entry_index,
                TestCfbDirectoryPatch {
                    start_sector: Some(word_entry.start_sector),
                    ..Default::default()
                },
            ),
            word_entry.entry_index,
            TestCfbDirectoryPatch {
                start_sector: Some(99),
                ..Default::default()
            },
        );

        let inspection = inspect_sectors_bytes(&patched).expect("inspection");
        let word = inspection
            .streams
            .iter()
            .find(|stream| stream.path == "WordDocument")
            .expect("word stream");
        let vba = inspection
            .streams
            .iter()
            .find(|stream| stream.path == "VBA/PROJECT")
            .expect("vba stream");
        assert_eq!(word.stream_health, "invalid-start");
        assert_eq!(word.stream_risk, "high");
        assert_eq!(vba.stream_health, "shared");
        assert_eq!(vba.stream_risk, "high");
        assert!(inspection.shared_chain_overlaps.iter().any(|overlap| {
            overlap.owners.iter().any(|owner| owner == "WordDocument") && overlap.severity == "high"
        }));
    }

    #[test]
    fn inspect_sectors_reports_mini_fat_incoherence_with_severity() {
        let base = build_test_cfb(&[("Tiny", &[0u8; 32])]);
        let patched =
            patch_test_cfb_header_u32(&patch_test_cfb_header_u32(&base, 0x3C, 0), 0x40, 1);

        let inspection = inspect_sectors_bytes(&patched).expect("inspection");
        assert!(inspection
            .anomalies
            .iter()
            .any(|entry| entry.kind == "mini-fat-without-consumers" && entry.severity == "medium"));
    }
}
