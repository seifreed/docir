use crate::{ArtifactInventory, ContainerKind, ParsedDocument};
use docir_core::security::ThreatLevel;

use super::DocumentIndicator;

pub(super) fn format_container_indicator(
    parsed: &ParsedDocument,
    inventory: &ArtifactInventory,
) -> DocumentIndicator {
    DocumentIndicator {
        key: "format-container".to_string(),
        value: format!(
            "{}/{}",
            parsed.format().extension(),
            container_label(inventory.container_kind)
        ),
        risk: ThreatLevel::None,
        reason: "Detected document format and source container".to_string(),
        evidence: Vec::new(),
    }
}

pub(super) fn boolean_or_count_indicator(
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

pub(super) fn container_label(kind: ContainerKind) -> &'static str {
    match kind {
        ContainerKind::ZipOoxml => "zip-ooxml",
        ContainerKind::ZipOdf => "zip-odf",
        ContainerKind::ZipHwpx => "zip-hwpx",
        ContainerKind::CfbOle => "cfb-ole",
        ContainerKind::Rtf => "rtf",
        ContainerKind::Unknown => "unknown",
    }
}
