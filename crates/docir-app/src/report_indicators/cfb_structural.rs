use crate::severity::severity_to_threat_level;
use docir_core::security::ThreatLevel;

use super::DocumentIndicator;
use super::helpers::boolean_or_count_indicator;

mod summary;
use summary::StructuralCfbSummary;

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
    let structural = summary::structural_cfb_summary(source_bytes);
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
