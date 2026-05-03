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
    if let Some(anomaly) = check_difat_mismatch(cfb) {
        anomalies.push(anomaly);
    }
    if let Some(anomaly) = check_fat_short(cfb, fat_entry_count) {
        anomalies.push(anomaly);
    }
    if let Some(anomaly) = check_mini_fat_without_consumers(cfb, streams) {
        anomalies.push(anomaly);
    }
    anomalies.extend(collect_shared_sector_anomalies(shared_sector_claims));
    anomalies.extend(collect_shared_chain_anomalies(shared_chain_overlaps));
    anomalies.extend(collect_start_sector_reuse_anomalies(start_sector_reuse));
    anomalies.extend(collect_structural_incoherence_anomalies(
        structural_incoherence_counts,
    ));
    anomalies
}

fn check_difat_mismatch(cfb: &docir_parser::ole::Cfb) -> Option<SectorAnomaly> {
    if cfb.num_fat_sectors() as usize > cfb.difat_entry_count() {
        Some(SectorAnomaly {
            kind: "difat-mismatch".to_string(),
            severity: "medium".to_string(),
            message: format!(
                "header declares {} FAT sector(s) but DIFAT exposes only {} entrie(s)",
                cfb.num_fat_sectors(),
                cfb.difat_entry_count()
            ),
        })
    } else {
        None
    }
}

fn check_fat_short(cfb: &docir_parser::ole::Cfb, fat_entry_count: usize) -> Option<SectorAnomaly> {
    if fat_entry_count < cfb.sector_count() as usize {
        Some(SectorAnomaly {
            kind: "fat-shorter-than-sector-count".to_string(),
            severity: "high".to_string(),
            message: format!(
                "fat entry table shorter than sector count ({} < {})",
                fat_entry_count,
                cfb.sector_count()
            ),
        })
    } else {
        None
    }
}

fn check_mini_fat_without_consumers(
    cfb: &docir_parser::ole::Cfb,
    streams: &[StreamSectorMap],
) -> Option<SectorAnomaly> {
    if cfb.mini_fat_entry_count() > 0
        && !streams.iter().any(|stream| stream.allocation == "mini-fat")
    {
        Some(SectorAnomaly {
            kind: "mini-fat-without-consumers".to_string(),
            severity: "medium".to_string(),
            message: "mini FAT is present but no stream currently resolves through mini-fat"
                .to_string(),
        })
    } else {
        None
    }
}

fn collect_shared_sector_anomalies(claims: &[SharedSectorClaim]) -> Vec<SectorAnomaly> {
    claims
        .iter()
        .map(|claim| SectorAnomaly {
            kind: "shared-sector".to_string(),
            severity: "high".to_string(),
            message: format!(
                "sector {} claimed by multiple streams: {}",
                claim.sector,
                claim.owners.join(", ")
            ),
        })
        .collect()
}

fn collect_shared_chain_anomalies(overlaps: &[SharedChainOverlap]) -> Vec<SectorAnomaly> {
    overlaps
        .iter()
        .map(|overlap| SectorAnomaly {
            kind: "shared-chain".to_string(),
            severity: "high".to_string(),
            message: format!(
                "streams {} share chain segment {:?}",
                overlap.owners.join(", "),
                overlap.sectors
            ),
        })
        .collect()
}

fn collect_start_sector_reuse_anomalies(reuse: &[StartSectorReuse]) -> Vec<SectorAnomaly> {
    reuse
        .iter()
        .map(|reuse| SectorAnomaly {
            kind: "start-sector-reuse".to_string(),
            severity: "medium".to_string(),
            message: format!(
                "start sector {} reused by streams: {}",
                reuse.sector,
                reuse.owners.join(", ")
            ),
        })
        .collect()
}

fn collect_structural_incoherence_anomalies(
    counts: &[StructuralIncoherenceCount],
) -> Vec<SectorAnomaly> {
    counts
        .iter()
        .map(|entry| SectorAnomaly {
            kind: entry.bucket.clone(),
            severity: entry.severity.clone(),
            message: format!("{}: {}", entry.bucket, entry.count),
        })
        .collect()
}
