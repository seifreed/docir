use docir_core::security::ThreatLevel;

/// Maps a threat-level string to a numeric rank for comparison.
pub fn severity_rank(value: &str) -> u8 {
    match value {
        "critical" => 4,
        "high" => 3,
        "medium" => 2,
        "low" => 1,
        _ => 0,
    }
}

/// Returns the canonical label for a numeric severity rank.
pub fn severity_label(rank: u8) -> &'static str {
    match rank {
        4 => "critical",
        3 => "high",
        2 => "medium",
        1 => "low",
        _ => "none",
    }
}

/// Returns the highest severity among the given string values.
pub fn max_severity<I>(values: I) -> String
where
    I: IntoIterator<Item = String>,
{
    let mut best_rank = 0u8;
    for value in values {
        let rank = severity_rank(&value);
        if rank > best_rank {
            best_rank = rank;
        }
    }
    severity_label(best_rank).to_string()
}

/// Converts a severity string to a ThreatLevel enum.
pub fn severity_to_threat_level(severity: &str) -> ThreatLevel {
    match severity {
        "critical" => ThreatLevel::Critical,
        "high" => ThreatLevel::High,
        "medium" => ThreatLevel::Medium,
        "low" => ThreatLevel::Low,
        _ => ThreatLevel::None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn severity_rank_maps_correctly() {
        assert_eq!(severity_rank("critical"), 4);
        assert_eq!(severity_rank("high"), 3);
        assert_eq!(severity_rank("medium"), 2);
        assert_eq!(severity_rank("low"), 1);
        assert_eq!(severity_rank("none"), 0);
        assert_eq!(severity_rank("unknown"), 0);
    }

    #[test]
    fn severity_label_round_trips() {
        for label in ["critical", "high", "medium", "low", "none"] {
            assert_eq!(severity_label(severity_rank(label)), label);
        }
    }

    #[test]
    fn max_severity_picks_highest() {
        assert_eq!(
            max_severity(vec![
                "low".to_string(),
                "high".to_string(),
                "medium".to_string()
            ]),
            "high"
        );
        assert_eq!(max_severity(vec![]), "none");
    }

    #[test]
    fn severity_to_threat_level_round_trips() {
        for (label, level) in [
            ("critical", ThreatLevel::Critical),
            ("high", ThreatLevel::High),
            ("medium", ThreatLevel::Medium),
            ("low", ThreatLevel::Low),
            ("none", ThreatLevel::None),
        ] {
            assert_eq!(severity_to_threat_level(label), level);
        }
    }
}
