use docir_core::security::ThreatLevel;
use serde::Serialize;

use crate::{ArtifactInventory, ParsedDocument, VbaRecognitionReport};

mod cfb_structural;
mod embedding_indicators;
mod helpers;
mod macro_indicators;

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
        indicators.push(helpers::format_container_indicator(parsed, &inventory));
        indicators.extend(macro_indicators::collect_macro_indicators(&vba));

        let embeddings = embedding_indicators::collect_embedding_indicators(parsed);
        indicators.push(helpers::boolean_or_count_indicator(
            "ole-objects",
            embeddings.ole_evidence.len(),
            ThreatLevel::High,
            "Embedded or linked OLE objects detected",
            "No OLE objects detected",
            embeddings.ole_evidence,
        ));
        indicators.push(helpers::boolean_or_count_indicator(
            "activex",
            embeddings.activex_evidence.len(),
            ThreatLevel::High,
            "ActiveX controls detected",
            "No ActiveX controls detected",
            embeddings.activex_evidence,
        ));
        indicators.push(helpers::boolean_or_count_indicator(
            "external-references",
            embeddings.external_reference_evidence.len(),
            ThreatLevel::Medium,
            "External references or remote relationships detected",
            "No external references detected",
            embeddings.external_reference_evidence,
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
        indicators.push(helpers::boolean_or_count_indicator(
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
        indicators.push(helpers::boolean_or_count_indicator(
            "object-pool",
            object_pool_evidence.len(),
            ThreatLevel::High,
            "Legacy ObjectPool storage entries detected",
            "No ObjectPool entries detected",
            object_pool_evidence,
        ));

        indicators.push(helpers::boolean_or_count_indicator(
            "native-payloads",
            embeddings.native_payload_evidence.len(),
            ThreatLevel::High,
            "Ole10Native or Package-style payloads detected",
            "No Ole10Native or Package-style payloads detected",
            embeddings.native_payload_evidence,
        ));

        indicators.push(embedding_indicators::collect_suspicious_relationships(
            security,
        ));

        if let Some(bytes) = source_bytes {
            indicators.extend(cfb_structural::collect_cfb_structural_indicators(bytes));
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
            container: helpers::container_label(inventory.container_kind).to_string(),
            overall_risk,
            indicator_count: indicators.len(),
            indicators,
        }
    }
}

#[cfg(test)]
mod tests;
