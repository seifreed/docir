use crate::types::NodeId;
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

pub(crate) fn max_indicator_threat_level(indicators: &[ThreatIndicator]) -> ThreatLevel {
    indicators
        .iter()
        .map(|indicator| indicator.severity)
        .max()
        .unwrap_or(ThreatLevel::None)
}

/// Threat level classification.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub enum ThreatLevel {
    /// No security concerns detected.
    #[default]
    None,
    /// Minor concerns (e.g., hyperlinks).
    Low,
    /// Moderate concerns (e.g., external templates).
    Medium,
    /// Significant concerns (e.g., OLE objects, DDE).
    High,
    /// Critical concerns (e.g., VBA macros with auto-exec).
    Critical,
}

impl std::fmt::Display for ThreatLevel {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "NONE"),
            Self::Low => write!(f, "LOW"),
            Self::Medium => write!(f, "MEDIUM"),
            Self::High => write!(f, "HIGH"),
            Self::Critical => write!(f, "CRITICAL"),
        }
    }
}

/// A specific threat indicator.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ThreatIndicator {
    /// Type of threat.
    pub indicator_type: ThreatIndicatorType,

    /// Severity level.
    pub severity: ThreatLevel,

    /// Human-readable description.
    pub description: String,

    /// Location in document (if applicable).
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub location: Option<String>,

    /// Related node ID (if applicable).
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub node_id: Option<NodeId>,
}

/// Types of threat indicators.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ThreatIndicatorType {
    /// VBA macro with auto-execute.
    AutoExecMacro,
    /// Macro project without auto-execute.
    MacroProject,
    /// Suspicious VBA API call.
    SuspiciousApiCall,
    /// External template reference.
    ExternalTemplate,
    /// Remote image/resource.
    RemoteResource,
    /// DDE command.
    DdeCommand,
    /// OLE object.
    OleObject,
    /// ActiveX control.
    ActiveXControl,
    /// Excel 4.0 XLM macro.
    XlmMacro,
    /// Hidden sheet with macros.
    HiddenMacroSheet,
    /// Suspicious formula.
    SuspiciousFormula,
    /// External hyperlink to suspicious domain.
    SuspiciousLink,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn max_indicator_threat_level_returns_highest() {
        let indicators = vec![
            ThreatIndicator {
                indicator_type: ThreatIndicatorType::ExternalTemplate,
                severity: ThreatLevel::Medium,
                description: "ext".to_string(),
                location: None,
                node_id: None,
            },
            ThreatIndicator {
                indicator_type: ThreatIndicatorType::SuspiciousApiCall,
                severity: ThreatLevel::Critical,
                description: "api".to_string(),
                location: None,
                node_id: None,
            },
        ];

        assert_eq!(
            max_indicator_threat_level(&indicators),
            ThreatLevel::Critical
        );
    }
}
