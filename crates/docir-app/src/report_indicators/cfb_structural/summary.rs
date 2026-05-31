use crate::{inspect_directory_bytes, inspect_sectors_bytes};

pub(super) struct StructuralCfbSummary {
    pub(super) all_evidence: Vec<String>,
    pub(super) shared_sector_evidence: Vec<String>,
    pub(super) live_unreachable_evidence: Vec<String>,
    pub(super) short_cycle_evidence: Vec<String>,
    pub(super) truncated_chain_evidence: Vec<String>,
    pub(super) chain_health_evidence: Vec<String>,
    pub(super) directory_score: String,
    pub(super) sector_score: String,
    pub(super) stream_score: String,
    pub(super) structural_score: String,
    pub(super) dominant_anomaly_class: String,
    pub(super) objectpool_corruption_evidence: Vec<String>,
    pub(super) vba_structure_anomalies_evidence: Vec<String>,
    pub(super) main_stream_corruption_evidence: Vec<String>,
}

pub(super) fn structural_cfb_summary(source_bytes: &[u8]) -> Option<StructuralCfbSummary> {
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
