use crate::severity::severity_to_threat_level;
use crate::{
    inspect_directory_bytes, inspect_sectors_bytes, ArtifactInventory, ContainerKind,
    ParsedDocument, VbaRecognitionReport,
};
use docir_core::ir::IRNode;
use docir_core::security::{ThreatIndicatorType, ThreatLevel};
use serde::Serialize;

/// Structured analyst-facing indicator summary for a parsed document.
#[derive(Debug, Clone, Serialize)]
pub struct IndicatorReport {
    pub document_format: String,
    pub container: String,
    pub overall_risk: ThreatLevel,
    pub indicator_count: usize,
    pub indicators: Vec<DocumentIndicator>,
}

/// A single document indicator with analyst-facing explanation and evidence.
#[derive(Debug, Clone, Serialize)]
pub struct DocumentIndicator {
    pub key: String,
    pub value: String,
    pub risk: ThreatLevel,
    pub reason: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub evidence: Vec<String>,
}

impl IndicatorReport {
    /// Builds a report from a parsed document using existing inventory and VBA recognition data.
    pub fn from_parsed(parsed: &ParsedDocument) -> Self {
        Self::from_parsed_with_bytes(parsed, None)
    }

    /// Builds a report from a parsed document and optional source bytes for low-level structure checks.
    pub fn from_parsed_with_bytes(parsed: &ParsedDocument, source_bytes: Option<&[u8]>) -> Self {
        let inventory = ArtifactInventory::from_parsed(parsed);
        let vba = VbaRecognitionReport::from_parsed(parsed, false);
        let security = parsed.security_info();

        let mut indicators = Vec::new();
        indicators.push(DocumentIndicator {
            key: "format-container".to_string(),
            value: format!(
                "{}/{}",
                parsed.format().extension(),
                container_label(inventory.container_kind)
            ),
            risk: ThreatLevel::None,
            reason: "Detected document format and source container".to_string(),
            evidence: Vec::new(),
        });

        let macro_count = vba.projects.len();
        let macro_evidence = vba
            .projects
            .iter()
            .map(|project| {
                project
                    .container_path
                    .clone()
                    .or_else(|| project.storage_root.clone())
                    .or_else(|| project.project_name.clone())
                    .unwrap_or_else(|| project.node_id.clone())
            })
            .collect::<Vec<_>>();
        indicators.push(boolean_or_count_indicator(
            "macros",
            macro_count,
            ThreatLevel::Critical,
            "VBA project or macro-capable content detected",
            "No macro-capable content detected",
            macro_evidence,
        ));

        let auto_exec = vba
            .projects
            .iter()
            .flat_map(|project| {
                project.auto_exec_procedures.iter().map(|proc_name| {
                    format!(
                        "{}:{}",
                        project
                            .project_name
                            .clone()
                            .unwrap_or_else(|| project.node_id.clone()),
                        proc_name
                    )
                })
            })
            .collect::<Vec<_>>();
        indicators.push(boolean_or_count_indicator(
            "autoexec",
            auto_exec.len(),
            ThreatLevel::Critical,
            "Auto-execute VBA entrypoints detected",
            "No auto-execute VBA entrypoints detected",
            auto_exec,
        ));

        let protected_projects = vba
            .projects
            .iter()
            .filter(|project| project.is_protected)
            .map(|project| {
                project
                    .project_name
                    .clone()
                    .or_else(|| project.storage_root.clone())
                    .unwrap_or_else(|| project.node_id.clone())
            })
            .collect::<Vec<_>>();
        indicators.push(boolean_or_count_indicator(
            "protected-vba",
            protected_projects.len(),
            ThreatLevel::Medium,
            "Password-protected VBA projects detected",
            "No password-protected VBA projects detected",
            protected_projects,
        ));

        let mut ole_evidence = Vec::new();
        let mut native_payload_evidence = Vec::new();
        let mut activex_evidence = Vec::new();
        let mut external_reference_evidence = Vec::new();
        for node in parsed.store().values() {
            match node {
                IRNode::OleObject(ole) => {
                    ole_evidence.push(ole_location_label(ole));
                    if ole.embedded_payload_kind.is_some()
                        || ole
                            .source_path
                            .as_deref()
                            .is_some_and(is_native_payload_path)
                    {
                        native_payload_evidence.push(ole_location_label(ole));
                    }
                }
                IRNode::ActiveXControl(control) => {
                    activex_evidence.push(
                        control
                            .span
                            .as_ref()
                            .map(|span| span.file_path.clone())
                            .or_else(|| control.name.clone())
                            .or_else(|| control.prog_id.clone())
                            .unwrap_or_else(|| control.id.to_string()),
                    );
                }
                IRNode::ExternalReference(reference) => {
                    external_reference_evidence
                        .push(format!("{:?}: {}", reference.ref_type, reference.target));
                }
                _ => {}
            }
        }

        indicators.push(boolean_or_count_indicator(
            "ole-objects",
            ole_evidence.len(),
            ThreatLevel::High,
            "Embedded or linked OLE objects detected",
            "No OLE objects detected",
            ole_evidence,
        ));
        indicators.push(boolean_or_count_indicator(
            "activex",
            activex_evidence.len(),
            ThreatLevel::High,
            "ActiveX controls detected",
            "No ActiveX controls detected",
            activex_evidence,
        ));
        indicators.push(boolean_or_count_indicator(
            "external-references",
            external_reference_evidence.len(),
            ThreatLevel::Medium,
            "External references or remote relationships detected",
            "No external references detected",
            external_reference_evidence,
        ));

        let dde_evidence = security
            .map(|info| {
                info.dde_fields
                    .iter()
                    .map(|dde| {
                        dde.location
                            .as_ref()
                            .map(|span| format!("{}: {}", span.file_path, dde.instruction))
                            .unwrap_or_else(|| dde.instruction.clone())
                    })
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        indicators.push(boolean_or_count_indicator(
            "dde",
            dde_evidence.len(),
            ThreatLevel::High,
            "Dynamic Data Exchange instructions detected",
            "No Dynamic Data Exchange instructions detected",
            dde_evidence,
        ));

        let object_pool_evidence = inventory
            .artifacts
            .iter()
            .filter_map(|artifact| artifact.path.as_ref())
            .filter(|path| path.contains("ObjectPool/"))
            .cloned()
            .collect::<Vec<_>>();
        indicators.push(boolean_or_count_indicator(
            "object-pool",
            object_pool_evidence.len(),
            ThreatLevel::High,
            "Legacy ObjectPool storage entries detected",
            "No ObjectPool entries detected",
            object_pool_evidence,
        ));

        indicators.push(boolean_or_count_indicator(
            "native-payloads",
            native_payload_evidence.len(),
            ThreatLevel::High,
            "Ole10Native or Package-style payloads detected",
            "No Ole10Native or Package-style payloads detected",
            native_payload_evidence,
        ));

        let suspicious_relationships = security
            .map(|info| {
                info.threat_indicators
                    .iter()
                    .filter(|indicator| {
                        matches!(
                            indicator.indicator_type,
                            ThreatIndicatorType::ExternalTemplate
                                | ThreatIndicatorType::RemoteResource
                                | ThreatIndicatorType::SuspiciousLink
                        )
                    })
                    .map(|indicator| {
                        indicator
                            .location
                            .as_ref()
                            .map(|location| format!("{location}: {}", indicator.description))
                            .unwrap_or_else(|| indicator.description.clone())
                    })
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        indicators.push(boolean_or_count_indicator(
            "suspicious-relationships",
            suspicious_relationships.len(),
            ThreatLevel::Medium,
            "Relationship-level external or suspicious targets detected",
            "No suspicious relationship targets detected",
            suspicious_relationships,
        ));

        if let Some(bytes) = source_bytes {
            let structural = structural_cfb_summary(bytes);
            let structural_evidence = structural
                .as_ref()
                .map(|summary| summary.all_evidence.clone())
                .unwrap_or_default();
            indicators.push(boolean_or_count_indicator(
                "cfb-structural-anomalies",
                structural_evidence.len(),
                ThreatLevel::High,
                "Low-level CFB structural anomalies detected",
                "No low-level CFB structural anomalies detected",
                structural_evidence,
            ));
            if let Some(summary) = structural {
                indicators.push(boolean_or_count_indicator(
                    "cfb-shared-sectors",
                    summary.shared_sector_evidence.len(),
                    ThreatLevel::High,
                    "Shared CFB sectors detected",
                    "No shared CFB sectors detected",
                    summary.shared_sector_evidence,
                ));
                indicators.push(boolean_or_count_indicator(
                    "cfb-live-unreachable",
                    summary.live_unreachable_evidence.len(),
                    ThreatLevel::High,
                    "Reachable graph misses live CFB entries",
                    "No unreachable live CFB entries detected",
                    summary.live_unreachable_evidence,
                ));
                indicators.push(boolean_or_count_indicator(
                    "cfb-short-cycles",
                    summary.short_cycle_evidence.len(),
                    ThreatLevel::High,
                    "Short directory cycles detected in CFB graph",
                    "No short directory cycles detected",
                    summary.short_cycle_evidence,
                ));
                indicators.push(boolean_or_count_indicator(
                    "cfb-truncated-chains",
                    summary.truncated_chain_evidence.len(),
                    ThreatLevel::High,
                    "Truncated CFB stream chains detected",
                    "No truncated CFB stream chains detected",
                    summary.truncated_chain_evidence,
                ));
                indicators.push(boolean_or_count_indicator(
                    "cfb-chain-health",
                    summary.chain_health_evidence.len(),
                    ThreatLevel::Medium,
                    "CFB chain health issues detected",
                    "No CFB chain health issues detected",
                    summary.chain_health_evidence.clone(),
                ));
                indicators.push(DocumentIndicator {
                    key: "cfb-directory-score".to_string(),
                    value: summary.directory_score.clone(),
                    risk: severity_to_threat_level(&summary.directory_score),
                    reason: "Aggregated directory-graph corruption score for the CFB container"
                        .to_string(),
                    evidence: summary
                        .all_evidence
                        .iter()
                        .filter(|entry| entry.starts_with("directory:"))
                        .cloned()
                        .collect(),
                });
                indicators.push(DocumentIndicator {
                    key: "cfb-sector-score".to_string(),
                    value: summary.sector_score.clone(),
                    risk: match summary.sector_score.as_str() {
                        "high" => ThreatLevel::High,
                        "medium" => ThreatLevel::Medium,
                        "low" => ThreatLevel::Low,
                        _ => ThreatLevel::None,
                    },
                    reason: "Aggregated sector-allocation corruption score for the CFB container"
                        .to_string(),
                    evidence: summary
                        .all_evidence
                        .iter()
                        .filter(|entry| entry.starts_with("sector:"))
                        .cloned()
                        .collect(),
                });
                indicators.push(DocumentIndicator {
                    key: "cfb-stream-score".to_string(),
                    value: summary.stream_score.clone(),
                    risk: match summary.stream_score.as_str() {
                        "high" => ThreatLevel::High,
                        "medium" => ThreatLevel::Medium,
                        "low" => ThreatLevel::Low,
                        _ => ThreatLevel::None,
                    },
                    reason: "Aggregated stream-level corruption score for the CFB container"
                        .to_string(),
                    evidence: summary.chain_health_evidence.clone(),
                });
                indicators.push(boolean_or_count_indicator(
                    "cfb-objectpool-corruption",
                    summary.objectpool_corruption_evidence.len(),
                    ThreatLevel::High,
                    "ObjectPool structural corruption detected",
                    "No ObjectPool structural corruption detected",
                    summary.objectpool_corruption_evidence,
                ));
                indicators.push(boolean_or_count_indicator(
                    "cfb-vba-structure-anomalies",
                    summary.vba_structure_anomalies_evidence.len(),
                    ThreatLevel::High,
                    "VBA-related structural anomalies detected",
                    "No VBA structural anomalies detected",
                    summary.vba_structure_anomalies_evidence,
                ));
                indicators.push(boolean_or_count_indicator(
                    "cfb-main-stream-corruption",
                    summary.main_stream_corruption_evidence.len(),
                    ThreatLevel::High,
                    "Main document stream corruption detected",
                    "No main document stream corruption detected",
                    summary.main_stream_corruption_evidence,
                ));
                indicators.push(DocumentIndicator {
                    key: "cfb-dominant-anomaly-class".to_string(),
                    value: summary.dominant_anomaly_class.clone(),
                    risk: if summary.dominant_anomaly_class == "none" {
                        ThreatLevel::None
                    } else {
                        ThreatLevel::Medium
                    },
                    reason: "Dominant low-level anomaly class across the CFB container".to_string(),
                    evidence: Vec::new(),
                });
                indicators.push(DocumentIndicator {
                    key: "cfb-structural-score".to_string(),
                    value: summary.structural_score.clone(),
                    risk: match summary.structural_score.as_str() {
                        "high" => ThreatLevel::High,
                        "medium" => ThreatLevel::Medium,
                        "low" => ThreatLevel::Low,
                        _ => ThreatLevel::None,
                    },
                    reason: "Aggregated structural corruption score for the CFB container"
                        .to_string(),
                    evidence: summary.all_evidence,
                });
            }
        }

        let overall_risk = security.map(|info| info.threat_level).unwrap_or_else(|| {
            indicators
                .iter()
                .map(|indicator| indicator.risk)
                .max()
                .unwrap_or_default()
        });

        Self {
            document_format: parsed.format().extension().to_string(),
            container: container_label(inventory.container_kind).to_string(),
            overall_risk,
            indicator_count: indicators.len(),
            indicators,
        }
    }
}

#[derive(Debug, Clone)]
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
    let mut all_evidence = Vec::new();
    let shared_sector_evidence = sectors
        .shared_sector_claims
        .iter()
        .map(|claim| format!("shared-sector:{}={}", claim.sector, claim.owners.join(",")))
        .collect::<Vec<_>>();
    let live_unreachable_evidence = directory
        .entries
        .iter()
        .filter(|entry| {
            entry.state == "normal"
                && entry.entry_type != "root-storage"
                && !entry.reachable_from_root
        })
        .map(|entry| format!("{} [{}]", entry.path, entry.anomaly_severity))
        .collect::<Vec<_>>();
    let short_cycle_evidence = directory
        .entries
        .iter()
        .flat_map(|entry| {
            entry
                .short_cycles
                .iter()
                .map(move |cycle| format!("{}:{cycle} [{}]", entry.path, entry.anomaly_severity))
        })
        .collect::<Vec<_>>();
    let truncated_chain_evidence = sectors
        .truncated_chain_counts
        .iter()
        .map(|entry| format!("{}={}", entry.bucket, entry.count))
        .collect::<Vec<_>>();
    let chain_health_evidence = sectors
        .chain_health_by_root
        .iter()
        .map(|entry| format!("{}={}", entry.bucket, entry.count))
        .collect::<Vec<_>>();
    let objectpool_corruption_evidence =
        collect_objectpool_corruption_evidence(&directory, &sectors);
    let vba_structure_anomalies_evidence =
        collect_vba_structure_anomalies_evidence(&directory, &sectors);
    let main_stream_corruption_evidence = collect_main_stream_corruption_evidence(&sectors);
    for entry in &directory.short_cycle_counts {
        all_evidence.push(format!("directory:cycle:{}={}", entry.bucket, entry.count));
    }
    for entry in &directory.reachability_counts {
        if entry.bucket == "live-unreachable" && entry.count > 0 {
            all_evidence.push(format!(
                "directory:reachability:{}={}",
                entry.bucket, entry.count
            ));
        }
    }
    for entry in &directory.incoming_source_counts {
        if entry.bucket == "incoming:state:anomalous" && entry.count > 0 {
            all_evidence.push(format!(
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
            all_evidence.push(format!(
                "directory:invalid-start:{} [{}]",
                entry.path, entry.anomaly_severity
            ));
        }
    }
    for claim in &sectors.shared_sector_claims {
        all_evidence.push(format!(
            "sector:shared-sector:{}={}",
            claim.sector,
            claim.owners.join(",")
        ));
    }
    for entry in &sectors.truncated_chain_counts {
        all_evidence.push(format!(
            "sector:truncated-chain:{}={}",
            entry.bucket, entry.count
        ));
    }
    for entry in &sectors.structural_incoherence_counts {
        all_evidence.push(format!(
            "sector:structural-incoherence:{}={} [{}]",
            entry.bucket, entry.count, entry.severity
        ));
    }
    let directory_score = directory.directory_score.clone();
    let sector_score = sectors.sector_score.clone();
    let stream_score = if sectors
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
    };
    let dominant_anomaly_class = dominant_anomaly_class(
        &shared_sector_evidence,
        &short_cycle_evidence,
        &live_unreachable_evidence,
        &sectors,
    );
    let structural_score = if sectors
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
    };
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

fn boolean_or_count_indicator(
    key: &str,
    count: usize,
    risk: ThreatLevel,
    present_reason: &str,
    absent_reason: &str,
    evidence: Vec<String>,
) -> DocumentIndicator {
    let (value, risk, reason) = if count > 0 {
        (count.to_string(), risk, present_reason.to_string())
    } else {
        (
            "absent".to_string(),
            ThreatLevel::None,
            absent_reason.to_string(),
        )
    };
    DocumentIndicator {
        key: key.to_string(),
        value,
        risk,
        reason,
        evidence,
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

fn container_label(kind: ContainerKind) -> &'static str {
    match kind {
        ContainerKind::ZipOoxml => "zip-ooxml",
        ContainerKind::ZipOdf => "zip-odf",
        ContainerKind::ZipHwpx => "zip-hwpx",
        ContainerKind::CfbOle => "cfb-ole",
        ContainerKind::Rtf => "rtf",
        ContainerKind::Unknown => "unknown",
    }
}

fn is_native_payload_path(path: &str) -> bool {
    path.ends_with("Ole10Native") || path.ends_with("Package")
}

fn ole_location_label(ole: &docir_core::security::OleObject) -> String {
    ole.source_path
        .clone()
        .or_else(|| ole.span.as_ref().map(|span| span.file_path.clone()))
        .or_else(|| ole.embedded_file_name.clone())
        .or_else(|| ole.class_name.clone())
        .or_else(|| ole.prog_id.clone())
        .unwrap_or_else(|| ole.id.to_string())
}

#[cfg(test)]
mod tests {
    use super::IndicatorReport;
    use crate::inspect_directory_bytes;
    use crate::test_support::{
        build_test_cfb, patch_test_cfb_directory_entry, patch_test_cfb_header_u32,
        TestCfbDirectoryPatch,
    };
    use crate::ParsedDocument;
    use docir_core::ir::{Document, IRNode};
    use docir_core::security::{
        ActiveXControl, DdeField, DdeFieldType, ExternalRefType, ExternalReference, MacroModule,
        MacroModuleType, MacroProject, ThreatIndicator, ThreatIndicatorType, ThreatLevel,
    };
    use docir_core::types::{DocumentFormat, SourceSpan};
    use docir_core::visitor::IrStore;

    fn make_parsed_with_security() -> ParsedDocument {
        let mut store = IrStore::new();

        let mut project = MacroProject::new();
        project.name = Some("LegacyProject".to_string());
        project.container_path = Some("word/vbaProject.bin".to_string());
        project.storage_root = Some("VBA".to_string());
        project.has_auto_exec = true;
        project.auto_exec_procedures = vec!["AutoOpen".to_string()];
        project.is_protected = true;

        let mut module = MacroModule::new("Module1", MacroModuleType::Standard);
        module.stream_path = Some("VBA/Module1".to_string());
        module.procedures = vec!["AutoOpen".to_string()];
        project.modules.push(module.id);

        let mut ole = docir_core::security::OleObject::new();
        ole.source_path = Some("ObjectPool/1/Ole10Native".to_string());
        ole.embedded_payload_kind = Some("ole10native".to_string());
        ole.size_bytes = 128;

        let mut activex = ActiveXControl::new();
        activex.name = Some("DangerousControl".to_string());

        let external = ExternalReference::new(
            ExternalRefType::AttachedTemplate,
            "https://evil.test/template.dotm",
        );

        let dde = DdeField {
            field_type: DdeFieldType::DdeAuto,
            application: "cmd".to_string(),
            topic: None,
            item: Some("/c calc".to_string()),
            instruction: r#"DDEAUTO "cmd" "/c calc""#.to_string(),
            location: Some(SourceSpan {
                file_path: "word/document.xml".to_string(),
                relationship_id: None,
                xml_path: None,
                line: None,
                column: None,
            }),
        };

        let mut doc = Document::new(DocumentFormat::WordProcessing);
        doc.security.macro_project = Some(project.id);
        doc.security.ole_objects.push(ole.id);
        doc.security.activex_controls.push(activex.id);
        doc.security.external_refs.push(external.id);
        doc.security.dde_fields.push(dde);
        doc.security.threat_level = ThreatLevel::Critical;
        doc.security.threat_indicators.push(ThreatIndicator {
            indicator_type: ThreatIndicatorType::ExternalTemplate,
            severity: ThreatLevel::Medium,
            description: "Remote template relationship".to_string(),
            location: Some("word/_rels/document.xml.rels".to_string()),
            node_id: None,
        });
        let root_id = doc.id;

        store.insert(IRNode::MacroProject(project));
        store.insert(IRNode::MacroModule(module));
        store.insert(IRNode::OleObject(ole));
        store.insert(IRNode::ActiveXControl(activex));
        store.insert(IRNode::ExternalReference(external));
        store.insert(IRNode::Document(doc));

        ParsedDocument::new(docir_parser::parser::ParsedDocument {
            root_id,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        })
    }

    #[test]
    fn report_indicators_collects_expected_signals() {
        let parsed = make_parsed_with_security();
        let report = IndicatorReport::from_parsed(&parsed);
        assert_eq!(report.container, "zip-ooxml");
        assert_eq!(report.overall_risk, ThreatLevel::Critical);
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "macros"
                && indicator.value == "1"
                && indicator.risk == ThreatLevel::Critical
        }));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "object-pool"
                && indicator
                    .evidence
                    .iter()
                    .any(|value| value.contains("ObjectPool/1/Ole10Native"))
        }));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "suspicious-relationships" && indicator.value == "1"
        }));
    }

    #[test]
    fn report_indicators_marks_absent_when_no_security_content_exists() {
        let mut store = IrStore::new();
        let doc = Document::new(DocumentFormat::WordProcessing);
        let root_id = doc.id;
        store.insert(IRNode::Document(doc));
        let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
            root_id,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        });

        let report = IndicatorReport::from_parsed(&parsed);
        assert_eq!(report.overall_risk, ThreatLevel::None);
        assert!(report
            .indicators
            .iter()
            .all(|indicator| indicator.key == "format-container" || indicator.value == "absent"));
    }

    #[test]
    fn report_indicators_includes_cfb_structural_anomalies_when_source_bytes_provided() {
        let mut store = IrStore::new();
        let doc = Document::new(DocumentFormat::WordProcessing);
        let root_id = doc.id;
        store.insert(IRNode::Document(doc));
        let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
            root_id,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        });

        let cfb = build_test_cfb(&[("WordDocument", b"doc")]);
        let patched = patch_test_cfb_header_u32(&patch_test_cfb_header_u32(&cfb, 0x3C, 0), 0x40, 1);
        let report = IndicatorReport::from_parsed_with_bytes(&parsed, Some(&patched));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "cfb-structural-anomalies"
                && indicator.value != "absent"
                && indicator
                    .evidence
                    .iter()
                    .any(|value| value.contains("structural-incoherence"))
        }));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "cfb-structural-score" && indicator.value == "medium"
        }));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "cfb-directory-score" && indicator.value == "none"
        }));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "cfb-sector-score" && indicator.value == "medium"
        }));
    }

    #[test]
    fn report_indicators_surface_specific_structural_classes() {
        let mut store = IrStore::new();
        let doc = Document::new(DocumentFormat::WordProcessing);
        let root_id = doc.id;
        store.insert(IRNode::Document(doc));
        let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
            root_id,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        });

        let base = build_test_cfb(&[
            ("WordDocument", b"doc"),
            ("VBA/PROJECT", b"meta"),
            ("VBA/Module1", b"code"),
            ("ObjectPool/1/Ole10Native", b"payload"),
        ]);
        let inspection = inspect_directory_bytes(&base).expect("directory");
        let word = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "WordDocument")
            .expect("word");
        let vba = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "VBA/PROJECT")
            .expect("vba");
        let objectpool = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "ObjectPool/1/Ole10Native")
            .expect("objectpool");

        let patched = patch_test_cfb_directory_entry(
            &patch_test_cfb_directory_entry(
                &patch_test_cfb_directory_entry(
                    &base,
                    vba.entry_index,
                    TestCfbDirectoryPatch {
                        start_sector: Some(word.start_sector),
                        ..Default::default()
                    },
                ),
                objectpool.entry_index,
                TestCfbDirectoryPatch {
                    start_sector: Some(99),
                    ..Default::default()
                },
            ),
            word.entry_index,
            TestCfbDirectoryPatch {
                start_sector: Some(98),
                ..Default::default()
            },
        );

        let report = IndicatorReport::from_parsed_with_bytes(&parsed, Some(&patched));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "cfb-objectpool-corruption" && indicator.value != "absent"
        }));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "cfb-vba-structure-anomalies" && indicator.value != "absent"
        }));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "cfb-main-stream-corruption" && indicator.value != "absent"
        }));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "cfb-directory-score" && indicator.value != "none"
        }));
        assert!(report
            .indicators
            .iter()
            .any(|indicator| { indicator.key == "cfb-sector-score" && indicator.value != "none" }));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "cfb-dominant-anomaly-class"
                && matches!(indicator.value.as_str(), "shared-sector" | "invalid-start")
        }));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "cfb-objectpool-corruption"
                && indicator
                    .evidence
                    .iter()
                    .any(|value| value.starts_with("objectpool:"))
        }));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "cfb-vba-structure-anomalies"
                && indicator
                    .evidence
                    .iter()
                    .any(|value| value.starts_with("vba:"))
        }));
        assert!(report.indicators.iter().any(|indicator| {
            indicator.key == "cfb-main-stream-corruption"
                && indicator
                    .evidence
                    .iter()
                    .any(|value| value.starts_with("main-stream:word:"))
        }));
    }

    #[test]
    fn dominant_anomaly_class_uses_stable_tie_break_order() {
        let mut store = IrStore::new();
        let doc = Document::new(DocumentFormat::WordProcessing);
        let root_id = doc.id;
        store.insert(IRNode::Document(doc));
        let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
            root_id,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        });
        let cfb = build_test_cfb(&[("WordDocument", b"doc")]);
        let patched = patch_test_cfb_header_u32(&patch_test_cfb_header_u32(&cfb, 0x3C, 0), 0x40, 1);
        let report = IndicatorReport::from_parsed_with_bytes(&parsed, Some(&patched));
        let dominant = report
            .indicators
            .iter()
            .find(|indicator| indicator.key == "cfb-dominant-anomaly-class")
            .expect("dominant anomaly indicator");
        assert_eq!(dominant.value, "mini-fat");
    }

    #[test]
    fn structural_indicator_taxonomy_prefixes_remain_stable() {
        let mut store = IrStore::new();
        let doc = Document::new(DocumentFormat::WordProcessing);
        let root_id = doc.id;
        store.insert(IRNode::Document(doc));
        let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
            root_id,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        });
        let base = build_test_cfb(&[
            ("WordDocument", b"doc"),
            ("VBA/PROJECT", b"vba"),
            ("ObjectPool/1/Ole10Native", b"obj"),
        ]);
        let inspection = inspect_directory_bytes(&base).expect("directory");
        let word = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "WordDocument")
            .expect("word");
        let vba = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "VBA/PROJECT")
            .expect("vba");
        let objectpool = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "ObjectPool/1/Ole10Native")
            .expect("objectpool");

        let patched = patch_test_cfb_directory_entry(
            &patch_test_cfb_directory_entry(
                &patch_test_cfb_directory_entry(
                    &base,
                    vba.entry_index,
                    TestCfbDirectoryPatch {
                        start_sector: Some(word.start_sector),
                        ..Default::default()
                    },
                ),
                objectpool.entry_index,
                TestCfbDirectoryPatch {
                    start_sector: Some(99),
                    ..Default::default()
                },
            ),
            word.entry_index,
            TestCfbDirectoryPatch {
                start_sector: Some(98),
                ..Default::default()
            },
        );
        let report = IndicatorReport::from_parsed_with_bytes(&parsed, Some(&patched));
        let evidence: Vec<&str> = report
            .indicators
            .iter()
            .flat_map(|indicator| indicator.evidence.iter().map(String::as_str))
            .collect();

        assert!(evidence.iter().any(|value| value.starts_with("directory:")));
        assert!(evidence.iter().any(|value| value.starts_with("sector:")));
        assert!(evidence.iter().any(|value| value.starts_with("health:")));
        assert!(evidence
            .iter()
            .any(|value| value.starts_with("objectpool:")));
        assert!(evidence.iter().any(|value| value.starts_with("vba:")));
        assert!(evidence
            .iter()
            .any(|value| value.starts_with("main-stream:")));
    }
}
