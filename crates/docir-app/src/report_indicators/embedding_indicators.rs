use crate::ParsedDocument;
use docir_core::ir::IRNode;
use docir_core::security::{ThreatIndicatorType, ThreatLevel};

use super::helpers::boolean_or_count_indicator;
use super::DocumentIndicator;

pub(super) struct EmbeddingIndicators {
    pub ole_evidence: Vec<String>,
    pub activex_evidence: Vec<String>,
    pub external_reference_evidence: Vec<String>,
    pub native_payload_evidence: Vec<String>,
}

pub(super) fn collect_embedding_indicators(parsed: &ParsedDocument) -> EmbeddingIndicators {
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
    EmbeddingIndicators {
        ole_evidence,
        activex_evidence,
        external_reference_evidence,
        native_payload_evidence,
    }
}

pub(super) fn collect_suspicious_relationships(
    security: Option<&docir_core::security::SecurityInfo>,
) -> DocumentIndicator {
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
    boolean_or_count_indicator(
        "suspicious-relationships",
        suspicious_relationships.len(),
        ThreatLevel::Medium,
        "Relationship-level external or suspicious targets detected",
        "No suspicious relationship targets detected",
        suspicious_relationships,
    )
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
