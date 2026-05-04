use crate::severity::severity_to_threat_level;
use crate::{inspect_directory_bytes, inspect_sectors_bytes};
use docir_core::security::ThreatLevel;

use super::helpers::boolean_or_count_indicator;
use super::DocumentIndicator;

fn score_indicator(
    key: &str,
    score: &str,
    reason: &str,
    evidence: Vec<String>,
) -> DocumentIndicator {
    DocumentIndicator {
        key: key.to_string(),
        value: score.to_string(),
        risk: severity_to_threat_level(score),
        reason: reason.to_string(),
        evidence,
    }
}

fn filtered_evidence(evidence: &[String], prefix: &str) -> Vec<String> {
    evidence
        .iter()
        .filter(|entry| entry.starts_with(prefix))
        .cloned()
        .collect()
}

pub(super) fn collect_cfb_structural_indicators(source_bytes: &[u8]) -> Vec<DocumentIndicator> {
    let structural = structural_cfb_summary(source_bytes);
    let structural_evidence = structural
        .as_ref()
        .map(|summary| summary.all_evidence.clone())
        .unwrap_or_default();
    let mut indicators = Vec::new();
    indicators.push(boolean_or_count_indicator(
        "cfb-structural-anomalies",
        structural_evidence.len(),
        ThreatLevel::High,
        "Low-level CFB structural anomalies detected",
        "No low-level CFB structural anomalies detected",
        structural_evidence,
    ));
    let Some(summary) = structural else {
        return indicators;
    };
    indicators.append(&mut summary_evidence_indicators(&summary));
    indicators.append(&mut summary_score_indicators(&summary));
    indicators.append(&mut summary_corruption_indicators(&summary));
    indicators
}

fn summary_evidence_indicators(summary: &StructuralCfbSummary) -> Vec<DocumentIndicator> {
    vec![
        boolean_or_count_indicator(
            "cfb-shared-sectors",
            summary.shared_sector_evidence.len(),
            ThreatLevel::High,
            "Shared CFB sectors detected",
            "No shared CFB sectors detected",
            summary.shared_sector_evidence.clone(),
        ),
        boolean_or_count_indicator(
            "cfb-live-unreachable",
            summary.live_unreachable_evidence.len(),
            ThreatLevel::High,
            "Reachable graph misses live CFB entries",
            "No unreachable live CFB entries detected",
            summary.live_unreachable_evidence.clone(),
        ),
        boolean_or_count_indicator(
            "cfb-short-cycles",
            summary.short_cycle_evidence.len(),
            ThreatLevel::High,
            "Short directory cycles detected in CFB graph",
            "No short directory cycles detected",
            summary.short_cycle_evidence.clone(),
        ),
        boolean_or_count_indicator(
            "cfb-truncated-chains",
            summary.truncated_chain_evidence.len(),
            ThreatLevel::High,
            "Truncated CFB stream chains detected",
            "No truncated CFB stream chains detected",
            summary.truncated_chain_evidence.clone(),
        ),
        boolean_or_count_indicator(
            "cfb-chain-health",
            summary.chain_health_evidence.len(),
            ThreatLevel::Medium,
            "CFB chain health issues detected",
            "No CFB chain health issues detected",
            summary.chain_health_evidence.clone(),
        ),
    ]
}

fn summary_score_indicators(summary: &StructuralCfbSummary) -> Vec<DocumentIndicator> {
    let all = &summary.all_evidence;
    vec![
        score_indicator(
            "cfb-directory-score",
            &summary.directory_score,
            "Aggregated directory-graph corruption score for the CFB container",
            filtered_evidence(all, "directory:"),
        ),
        score_indicator(
            "cfb-sector-score",
            &summary.sector_score,
            "Aggregated sector-allocation corruption score for the CFB container",
            filtered_evidence(all, "sector:"),
        ),
        score_indicator(
            "cfb-stream-score",
            &summary.stream_score,
            "Aggregated stream-level corruption score for the CFB container",
            summary.chain_health_evidence.clone(),
        ),
        DocumentIndicator {
            key: "cfb-dominant-anomaly-class".to_string(),
            value: summary.dominant_anomaly_class.clone(),
            risk: if summary.dominant_anomaly_class == "none" {
                ThreatLevel::None
            } else {
                ThreatLevel::Medium
            },
            reason: "Dominant low-level anomaly class across the CFB container".to_string(),
            evidence: Vec::new(),
        },
        score_indicator(
            "cfb-structural-score",
            &summary.structural_score,
            "Aggregated structural corruption score for the CFB container",
            all.clone(),
        ),
    ]
}

fn summary_corruption_indicators(summary: &StructuralCfbSummary) -> Vec<DocumentIndicator> {
    vec![
        boolean_or_count_indicator(
            "cfb-objectpool-corruption",
            summary.objectpool_corruption_evidence.len(),
            ThreatLevel::High,
            "ObjectPool structural corruption detected",
            "No ObjectPool structural corruption detected",
            summary.objectpool_corruption_evidence.clone(),
        ),
        boolean_or_count_indicator(
            "cfb-vba-structure-anomalies",
            summary.vba_structure_anomalies_evidence.len(),
            ThreatLevel::High,
            "VBA-related structural anomalies detected",
            "No VBA structural anomalies detected",
            summary.vba_structure_anomalies_evidence.clone(),
        ),
        boolean_or_count_indicator(
            "cfb-main-stream-corruption",
            summary.main_stream_corruption_evidence.len(),
            ThreatLevel::High,
            "Main document stream corruption detected",
            "No main document stream corruption detected",
            summary.main_stream_corruption_evidence.clone(),
        ),
    ]
}

struct StructuralCfbSummary {
    all_evidence: Vec<String>,
    shared_sector_evidence: Vec<String>,
    live_unreachable_evidence: Vec<String>,
    short_cycle_evidence: Vec<String>,
    truncated_chain_evidence: Vec<String>,
    chain_health_evidence: Vec<String>,
    directory_score: String,
    sector_score: String,
    stream_score: String,
    structural_score: String,
    dominant_anomaly_class: String,
    objectpool_corruption_evidence: Vec<String>,
    vba_structure_anomalies_evidence: Vec<String>,
    main_stream_corruption_evidence: Vec<String>,
}

fn structural_cfb_summary(source_bytes: &[u8]) -> Option<StructuralCfbSummary> {
    let directory = inspect_directory_bytes(source_bytes).ok()?;
    let sectors = inspect_sectors_bytes(source_bytes).ok()?;

    let shared_sector_evidence = collect_shared_sector_evidence(&sectors);
    let live_unreachable_evidence = collect_live_unreachable_evidence(&directory);
    let short_cycle_evidence = collect_short_cycle_evidence(&directory);
    let truncated_chain_evidence = collect_truncated_chain_evidence(&sectors);
    let chain_health_evidence = collect_chain_health_evidence(&sectors);
    let objectpool_corruption_evidence =
        collect_objectpool_corruption_evidence(&directory, &sectors);
    let vba_structure_anomalies_evidence =
        collect_vba_structure_anomalies_evidence(&directory, &sectors);
    let main_stream_corruption_evidence = collect_main_stream_corruption_evidence(&sectors);
    let all_evidence = collect_all_evidence(&directory, &sectors);

    let directory_score = directory.directory_score.clone();
    let sector_score = sectors.sector_score.clone();
    let stream_score = compute_stream_score(&sectors);
    let dominant_anomaly_class = dominant_anomaly_class(
        &shared_sector_evidence,
        &short_cycle_evidence,
        &live_unreachable_evidence,
        &sectors,
    );
    let structural_score = compute_structural_score(
        &sectors,
        &shared_sector_evidence,
        &short_cycle_evidence,
        &live_unreachable_evidence,
        &truncated_chain_evidence,
        &all_evidence,
    );

    Some(StructuralCfbSummary {
        all_evidence,
        shared_sector_evidence,
        live_unreachable_evidence,
        short_cycle_evidence,
        truncated_chain_evidence,
        chain_health_evidence,
        directory_score,
        sector_score,
        stream_score: stream_score.to_string(),
        structural_score: structural_score.to_string(),
        dominant_anomaly_class: dominant_anomaly_class.to_string(),
        objectpool_corruption_evidence,
        vba_structure_anomalies_evidence,
        main_stream_corruption_evidence,
    })
}

fn collect_shared_sector_evidence(
    sectors: &crate::inspect_sectors::SectorInspection,
) -> Vec<String> {
    sectors
        .shared_sector_claims
        .iter()
        .map(|claim| format!("shared-sector:{}={}", claim.sector, claim.owners.join(",")))
        .collect()
}

fn collect_live_unreachable_evidence(
    directory: &crate::inspect_directory::DirectoryInspection,
) -> Vec<String> {
    directory
        .entries
        .iter()
        .filter(|entry| {
            entry.state == "normal"
                && entry.entry_type != "root-storage"
                && !entry.reachable_from_root
        })
        .map(|entry| format!("{} [{}]", entry.path, entry.anomaly_severity))
        .collect()
}

fn collect_short_cycle_evidence(
    directory: &crate::inspect_directory::DirectoryInspection,
) -> Vec<String> {
    directory
        .entries
        .iter()
        .flat_map(|entry| {
            entry
                .short_cycles
                .iter()
                .map(move |cycle| format!("{}:{cycle} [{}]", entry.path, entry.anomaly_severity))
        })
        .collect()
}

fn collect_truncated_chain_evidence(
    sectors: &crate::inspect_sectors::SectorInspection,
) -> Vec<String> {
    sectors
        .truncated_chain_counts
        .iter()
        .map(|entry| format!("{}={}", entry.bucket, entry.count))
        .collect()
}

fn collect_chain_health_evidence(
    sectors: &crate::inspect_sectors::SectorInspection,
) -> Vec<String> {
    sectors
        .chain_health_by_root
        .iter()
        .map(|entry| format!("{}={}", entry.bucket, entry.count))
        .collect()
}

fn collect_all_evidence(
    directory: &crate::inspect_directory::DirectoryInspection,
    sectors: &crate::inspect_sectors::SectorInspection,
) -> Vec<String> {
    let mut evidence = Vec::new();
    for entry in &directory.short_cycle_counts {
        evidence.push(format!("directory:cycle:{}={}", entry.bucket, entry.count));
    }
    for entry in &directory.reachability_counts {
        if entry.bucket == "live-unreachable" && entry.count > 0 {
            evidence.push(format!(
                "directory:reachability:{}={}",
                entry.bucket, entry.count
            ));
        }
    }
    for entry in &directory.incoming_source_counts {
        if entry.bucket == "incoming:state:anomalous" && entry.count > 0 {
            evidence.push(format!(
                "directory:incoming:{}={}",
                entry.bucket, entry.count
            ));
        }
    }
    for entry in &directory.entries {
        if entry
            .anomaly_tags
            .iter()
            .any(|tag| tag == "invalid-start-sector")
        {
            evidence.push(format!(
                "directory:invalid-start:{} [{}]",
                entry.path, entry.anomaly_severity
            ));
        }
    }
    for claim in &sectors.shared_sector_claims {
        evidence.push(format!(
            "sector:shared-sector:{}={}",
            claim.sector,
            claim.owners.join(",")
        ));
    }
    for entry in &sectors.truncated_chain_counts {
        evidence.push(format!(
            "sector:truncated-chain:{}={}",
            entry.bucket, entry.count
        ));
    }
    for entry in &sectors.structural_incoherence_counts {
        evidence.push(format!(
            "sector:structural-incoherence:{}={} [{}]",
            entry.bucket, entry.count, entry.severity
        ));
    }
    evidence
}

fn compute_stream_score(sectors: &crate::inspect_sectors::SectorInspection) -> &'static str {
    if sectors
        .streams
        .iter()
        .any(|stream| stream.stream_risk == "high")
    {
        "high"
    } else if sectors
        .streams
        .iter()
        .any(|stream| stream.stream_risk == "medium")
    {
        "medium"
    } else if sectors
        .streams
        .iter()
        .any(|stream| stream.stream_risk == "low")
    {
        "low"
    } else {
        "none"
    }
}

fn compute_structural_score(
    sectors: &crate::inspect_sectors::SectorInspection,
    shared_sector_evidence: &[String],
    short_cycle_evidence: &[String],
    live_unreachable_evidence: &[String],
    truncated_chain_evidence: &[String],
    all_evidence: &[String],
) -> &'static str {
    if sectors
        .anomalies
        .iter()
        .any(|anomaly| anomaly.severity == "high")
        || !shared_sector_evidence.is_empty()
        || !short_cycle_evidence.is_empty()
    {
        "high"
    } else if !live_unreachable_evidence.is_empty()
        || !truncated_chain_evidence.is_empty()
        || sectors
            .anomalies
            .iter()
            .any(|anomaly| anomaly.severity == "medium")
    {
        "medium"
    } else if !all_evidence.is_empty() {
        "low"
    } else {
        "none"
    }
}

fn collect_objectpool_corruption_evidence(
    directory: &crate::inspect_directory::DirectoryInspection,
    sectors: &crate::inspect_sectors::SectorInspection,
) -> Vec<String> {
    directory
        .entries
        .iter()
        .filter(|entry| entry.path.contains("ObjectPool/") && entry.anomaly_severity != "none")
        .flat_map(|entry| {
            entry.anomaly_tags.iter().map(move |tag| {
                format!(
                    "objectpool:{}:{} [{}]",
                    normalize_objectpool_bucket(tag),
                    entry.path,
                    entry.anomaly_severity
                )
            })
        })
        .chain(
            sectors
                .streams
                .iter()
                .filter(|stream| {
                    stream.path.contains("ObjectPool/") && stream.stream_risk != "none"
                })
                .map(|stream| {
                    format!(
                        "objectpool:{}:{} [{}:{}]",
                        normalize_stream_health_bucket(&stream.stream_health),
                        stream.path,
                        stream.stream_health,
                        stream.stream_risk
                    )
                }),
        )
        .collect()
}

fn collect_vba_structure_anomalies_evidence(
    directory: &crate::inspect_directory::DirectoryInspection,
    sectors: &crate::inspect_sectors::SectorInspection,
) -> Vec<String> {
    directory
        .entries
        .iter()
        .filter(|entry| {
            (entry.path == "VBA" || entry.path.starts_with("VBA/"))
                && entry.anomaly_severity != "none"
        })
        .map(|entry| {
            let kind = if entry.path == "VBA" {
                "vba:storage"
            } else if entry.path == "VBA/PROJECT" {
                "vba:project-stream"
            } else {
                "vba:module-stream"
            };
            format!("{kind}:{} [{}]", entry.path, entry.anomaly_severity)
        })
        .chain(
            sectors
                .streams
                .iter()
                .filter(|stream| {
                    (stream.path == "VBA" || stream.path.starts_with("VBA/"))
                        && stream.stream_risk != "none"
                })
                .map(|stream| {
                    let kind = if stream.path == "VBA/PROJECT" {
                        "vba:project-stream"
                    } else if stream.path.starts_with("VBA/") {
                        "vba:module-stream"
                    } else {
                        "vba:storage"
                    };
                    format!(
                        "{kind}:{} [{}:{}]",
                        stream.path, stream.stream_health, stream.stream_risk
                    )
                }),
        )
        .collect()
}

fn collect_main_stream_corruption_evidence(
    sectors: &crate::inspect_sectors::SectorInspection,
) -> Vec<String> {
    sectors
        .streams
        .iter()
        .filter_map(|stream| {
            if stream.stream_risk == "none" {
                return None;
            }
            let bucket = match stream.logical_root.as_str() {
                "WordDocument" => Some("main-stream:word"),
                "Workbook" => Some("main-stream:xls"),
                "PowerPoint Document" => Some("main-stream:ppt"),
                _ => None,
            }?;
            Some(format!(
                "{bucket}:{} [{}:{}]",
                stream.path, stream.stream_health, stream.stream_risk
            ))
        })
        .collect()
}

fn normalize_objectpool_bucket(tag: &str) -> &'static str {
    if tag.contains("orphaned") {
        "orphaned"
    } else if tag.contains("shared") {
        "shared"
    } else if tag.contains("invalid-start") {
        "invalid-start"
    } else if tag.contains("truncated") {
        "truncated"
    } else {
        "anomalous"
    }
}

fn normalize_stream_health_bucket(health: &str) -> &'static str {
    match health {
        "shared" => "shared",
        "invalid-start" => "invalid-start",
        "truncated" => "truncated",
        "start-reused" => "shared",
        _ => "anomalous",
    }
}

fn dominant_anomaly_class(
    shared_sector_evidence: &[String],
    short_cycle_evidence: &[String],
    live_unreachable_evidence: &[String],
    sectors: &crate::inspect_sectors::SectorInspection,
) -> &'static str {
    let shared_score = (shared_sector_evidence.len() * 4)
        + sectors
            .shared_chain_overlaps
            .iter()
            .map(|entry| if entry.severity == "high" { 3 } else { 2 })
            .sum::<usize>();
    let cycle_score = short_cycle_evidence.len() * 3;
    let unreachable_score = live_unreachable_evidence.len() * 3;
    let invalid_start_score = sectors
        .streams
        .iter()
        .filter(|stream| stream.stream_health == "invalid-start")
        .count()
        * 4;
    let mini_fat_score = sectors
        .structural_incoherence_counts
        .iter()
        .filter(|entry| entry.bucket == "mini-fat-without-consumers")
        .map(|entry| entry.count * 2)
        .sum::<usize>();

    let classes = [
        ("shared-sector", shared_score),
        ("cycle", cycle_score),
        ("unreachable-live", unreachable_score),
        ("invalid-start", invalid_start_score),
        ("mini-fat", mini_fat_score),
    ];
    classes
        .into_iter()
        .enumerate()
        .max_by_key(|(index, (_, score))| (*score, usize::MAX - *index))
        .map(|(_, (class, score))| (class, score))
        .and_then(|(class, score)| if score > 0 { Some(class) } else { None })
        .unwrap_or("none")
}
