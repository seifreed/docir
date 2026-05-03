use serde::Serialize;

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
