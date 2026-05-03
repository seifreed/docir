use super::types::*;

pub(super) fn collect_cfb_anomalies(
    cfb: &docir_parser::ole::Cfb,
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
