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

impl std::str::FromStr for ThreatLevel {
    type Err = ParseThreatLevelError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" | "NONE" => Ok(Self::None),
            "low" | "LOW" => Ok(Self::Low),
            "medium" | "MEDIUM" => Ok(Self::Medium),
            "high" | "HIGH" => Ok(Self::High),
            "critical" | "CRITICAL" => Ok(Self::Critical),
            _ => Err(ParseThreatLevelError {
                input: s.to_string(),
            }),
        }
    }
}

/// Error returned when a string cannot be parsed as a [`ThreatLevel`].
#[derive(Debug, Clone, PartialEq)]
pub struct ParseThreatLevelError {
    input: String,
}

impl std::fmt::Display for ParseThreatLevelError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "invalid threat level: {:?} (expected none/low/medium/high/critical)",
            self.input
        )
    }
}

impl std::error::Error for ParseThreatLevelError {}

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

    #[test]
    fn display_produces_uppercase() {
        assert_eq!(ThreatLevel::None.to_string(), "NONE");
        assert_eq!(ThreatLevel::Low.to_string(), "LOW");
        assert_eq!(ThreatLevel::Medium.to_string(), "MEDIUM");
        assert_eq!(ThreatLevel::High.to_string(), "HIGH");
        assert_eq!(ThreatLevel::Critical.to_string(), "CRITICAL");
    }

    #[test]
    fn from_str_parses_lowercase() {
        assert_eq!("critical".parse::<ThreatLevel>(), Ok(ThreatLevel::Critical));
        assert_eq!("high".parse::<ThreatLevel>(), Ok(ThreatLevel::High));
        assert_eq!("medium".parse::<ThreatLevel>(), Ok(ThreatLevel::Medium));
        assert_eq!("low".parse::<ThreatLevel>(), Ok(ThreatLevel::Low));
        assert_eq!("none".parse::<ThreatLevel>(), Ok(ThreatLevel::None));
    }

    #[test]
    fn from_str_parses_uppercase() {
        assert_eq!("CRITICAL".parse::<ThreatLevel>(), Ok(ThreatLevel::Critical));
        assert_eq!("HIGH".parse::<ThreatLevel>(), Ok(ThreatLevel::High));
        assert_eq!("MEDIUM".parse::<ThreatLevel>(), Ok(ThreatLevel::Medium));
        assert_eq!("LOW".parse::<ThreatLevel>(), Ok(ThreatLevel::Low));
        assert_eq!("NONE".parse::<ThreatLevel>(), Ok(ThreatLevel::None));
    }

    #[test]
    fn from_str_rejects_unknown() {
        assert!("unknown".parse::<ThreatLevel>().is_err());
        assert!("".parse::<ThreatLevel>().is_err());
    }

    #[test]
    fn display_from_str_round_trip() {
        for level in [
            ThreatLevel::None,
            ThreatLevel::Low,
            ThreatLevel::Medium,
            ThreatLevel::High,
            ThreatLevel::Critical,
        ] {
            assert_eq!(level.to_string().parse::<ThreatLevel>(), Ok(level));
        }
    }
}
