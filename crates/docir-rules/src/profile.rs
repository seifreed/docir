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
