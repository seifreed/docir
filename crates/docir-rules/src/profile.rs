//! Rule profile configuration and filtering helpers.

use crate::Rule;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, HashSet};

/// Rule profile configuration.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RuleProfile {
    #[serde(default)]
    pub enabled_rules: Option<Vec<String>>,
    #[serde(default)]
    pub disabled_rules: Vec<String>,
    #[serde(default)]
    pub severity_overrides: BTreeMap<String, crate::Severity>,
    #[serde(default)]
    pub thresholds: RuleThresholds,
}

impl Default for RuleProfile {
    fn default() -> Self {
        Self {
            enabled_rules: None,
            disabled_rules: Vec::new(),
            severity_overrides: BTreeMap::new(),
            thresholds: RuleThresholds::default(),
        }
    }
}

/// Thresholds for rule tuning.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RuleThresholds {
    #[serde(default)]
    pub max_external_links: Option<usize>,
    #[serde(default)]
    pub max_ole_objects: Option<usize>,
    #[serde(default)]
    pub max_activex_controls: Option<usize>,
}

impl Default for RuleThresholds {
    fn default() -> Self {
        Self {
            max_external_links: None,
            max_ole_objects: None,
            max_activex_controls: None,
        }
    }
}

pub(crate) fn apply_profile(
    rules: Vec<Box<dyn Rule>>,
    profile: &RuleProfile,
) -> Vec<Box<dyn Rule>> {
    if profile.enabled_rules.is_none() && profile.disabled_rules.is_empty() {
        return rules;
    }
    let enabled = profile
        .enabled_rules
        .as_ref()
        .map(|list| list.iter().cloned().collect::<HashSet<_>>());
    let disabled: HashSet<String> = profile.disabled_rules.iter().cloned().collect();
    rules
        .into_iter()
        .filter(|rule| {
            if disabled.contains(rule.id()) {
                return false;
            }
            if let Some(set) = &enabled {
                return set.contains(rule.id());
            }
            true
        })
        .collect()
}

pub(crate) fn profile_rule_enabled(profile: &RuleProfile, rule_id: &str) -> bool {
    if profile.disabled_rules.iter().any(|r| r == rule_id) {
        return false;
    }
    if let Some(enabled) = &profile.enabled_rules {
        return enabled.iter().any(|r| r == rule_id);
    }
    true
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Finding, Rule, RuleCategory, RuleContext, Severity};

    struct DummyRule {
        id: &'static str,
    }

    impl Rule for DummyRule {
        fn id(&self) -> &'static str {
            self.id
        }

        fn name(&self) -> &'static str {
            self.id
        }

        fn description(&self) -> &'static str {
            "dummy"
        }

        fn category(&self) -> RuleCategory {
            RuleCategory::Security
        }

        fn default_severity(&self) -> Severity {
            Severity::Low
        }

        fn run(&self, _ctx: &RuleContext, _findings: &mut Vec<Finding>) {}
    }

    #[test]
    fn apply_profile_returns_all_rules_when_profile_has_no_filters() {
        let rules: Vec<Box<dyn Rule>> = vec![
            Box::new(DummyRule { id: "a" }),
            Box::new(DummyRule { id: "b" }),
        ];
        let profile = RuleProfile::default();
        let filtered = apply_profile(rules, &profile);
        assert_eq!(filtered.len(), 2);
    }

    #[test]
    fn apply_profile_respects_enabled_and_disabled_rules() {
        let rules: Vec<Box<dyn Rule>> = vec![
            Box::new(DummyRule { id: "a" }),
            Box::new(DummyRule { id: "b" }),
            Box::new(DummyRule { id: "c" }),
        ];
        let profile = RuleProfile {
            enabled_rules: Some(vec!["a".to_string(), "b".to_string()]),
            disabled_rules: vec!["b".to_string()],
            severity_overrides: BTreeMap::new(),
            thresholds: RuleThresholds::default(),
        };

        let filtered = apply_profile(rules, &profile);
        let ids: Vec<_> = filtered.into_iter().map(|r| r.id().to_string()).collect();
        assert_eq!(ids, vec!["a".to_string()]);
    }

    #[test]
    fn profile_rule_enabled_prioritizes_disabled_list() {
        let profile = RuleProfile {
            enabled_rules: Some(vec!["r1".to_string(), "r2".to_string()]),
            disabled_rules: vec!["r2".to_string(), "r3".to_string()],
            severity_overrides: BTreeMap::new(),
            thresholds: RuleThresholds::default(),
        };

        assert!(profile_rule_enabled(&profile, "r1"));
        assert!(!profile_rule_enabled(&profile, "r2"));
        assert!(!profile_rule_enabled(&profile, "r3"));
        assert!(!profile_rule_enabled(&profile, "other"));

        let open_profile = RuleProfile::default();
        assert!(profile_rule_enabled(&open_profile, "any-rule"));
    }
}
