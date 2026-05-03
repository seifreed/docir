use super::types::*;
use crate::severity::max_severity;

pub(super) fn stream_health_tags(
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

pub(super) fn build_chain_health_counts_by_root(
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

pub(super) fn build_chain_health_counts_by_allocation(
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

pub(super) fn build_sector_score(
    streams: &[StreamSectorMap],
    anomalies: &[SectorAnomaly],
) -> String {
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

pub(super) fn apply_stream_health(
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
