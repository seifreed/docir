use super::types::*;

pub(super) fn build_role_counts(sector_overview: &[SectorOverviewEntry]) -> Vec<RoleCount> {
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

pub(super) fn build_shared_sector_claims(
    sector_overview: &[SectorOverviewEntry],
) -> Vec<SharedSectorClaim> {
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

pub(super) fn build_shared_chain_overlaps(
    sector_overview: &[SectorOverviewEntry],
) -> Vec<SharedChainOverlap> {
    let mut grouped = std::collections::BTreeMap::<Vec<String>, Vec<(u32, Vec<usize>)>>::new();
    for entry in sector_overview
        .iter()
        .filter(|entry| entry.owners.len() > 1)
    {
        let mut owners = entry
            .owners
            .iter()
            .map(|owner| (owner.path.clone(), owner.index_in_chain))
            .collect::<Vec<_>>();
        owners.sort_unstable_by(|left, right| left.0.cmp(&right.0));
        let paths = owners.iter().map(|(path, _)| path.clone()).collect();
        let positions = owners.into_iter().map(|(_, index)| index).collect();
        grouped
            .entry(paths)
            .or_default()
            .push((entry.sector, positions));
    }

    let mut overlaps = Vec::new();
    for (owners, mut claims) in grouped {
        claims.sort_unstable_by_key(|(_, positions)| positions[0]);
        let mut run = Vec::new();
        let mut previous_positions: Option<Vec<usize>> = None;
        for (sector, positions) in claims {
            let is_next = previous_positions.as_ref().is_some_and(|previous| {
                positions
                    .iter()
                    .zip(previous)
                    .all(|(current, previous)| *current == previous.saturating_add(1))
            });
            if !is_next && run.len() >= 2 {
                let severity = shared_chain_severity(&owners, &run).to_string();
                overlaps.push(SharedChainOverlap {
                    owners: owners.clone(),
                    sectors: run,
                    severity,
                });
                run = Vec::new();
            }
            run.push(sector);
            previous_positions = Some(positions);
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

pub(super) fn build_start_sector_reuse(streams: &[StreamSectorMap]) -> Vec<StartSectorReuse> {
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

pub(super) fn build_truncated_chain_counts(
    streams: &[StreamSectorMap],
) -> Vec<TruncatedChainCount> {
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

pub(super) fn build_structural_incoherence_counts(
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
