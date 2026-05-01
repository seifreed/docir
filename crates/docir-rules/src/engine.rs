//! Rule engine core types and execution.

use crate::profile::{apply_profile, profile_rule_enabled, RuleProfile, RuleThresholds};
use crate::rules::default_rules;
use docir_core::ir::{Document, DocumentMetadata, IRNode};
use docir_core::types::{NodeId, NodeType};
use docir_core::visitor::IrStore;
use serde::{Deserialize, Serialize};

/// Severity level for rule findings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, PartialOrd, Ord)]
pub enum Severity {
    Info,
    Low,
    Medium,
    High,
    Critical,
}

/// Rule category.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum RuleCategory {
    Security,
    Structure,
    Content,
    Metadata,
}

/// Rule finding produced by the engine.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Finding {
    pub rule_id: String,
    pub rule_name: String,
    pub category: RuleCategory,
    pub severity: Severity,
    pub message: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub context: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub node_id: Option<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub node_type: Option<NodeType>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub location: Option<String>,
}

/// Rule execution result.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RuleReport {
    pub findings: Vec<Finding>,
}

impl RuleReport {
    /// Public API entrypoint: is_empty.
    pub fn is_empty(&self) -> bool {
        self.findings.is_empty()
    }
}

/// Rule trait.
pub trait Rule: Send + Sync {
    fn id(&self) -> &'static str;
    fn name(&self) -> &'static str;
    fn description(&self) -> &'static str;
    fn category(&self) -> RuleCategory;
    fn default_severity(&self) -> Severity;
    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>);
}

/// Rule execution context with document-level enrichment.
pub struct RuleContext<'a> {
    pub store: &'a IrStore,
    pub root: NodeId,
    pub document: Option<&'a Document>,
    pub metadata: Option<&'a DocumentMetadata>,
    pub thresholds: RuleThresholds,
}

/// Rule engine.
pub struct RuleEngine {
    rules: Vec<Box<dyn Rule>>,
}

impl RuleEngine {
    /// Creates a new engine with default rules.
    pub fn with_default_rules() -> Self {
        Self {
            rules: default_rules(),
        }
    }

    /// Creates a new engine with default rules and a profile.
    pub fn with_profile(profile: RuleProfile) -> Self {
        let mut engine = Self::with_default_rules();
        engine.rules = apply_profile(engine.rules, &profile);
        engine
    }

    /// Adds a custom rule.
    pub fn add_rule(&mut self, rule: Box<dyn Rule>) {
        self.rules.push(rule);
    }

    /// Runs the rules and returns findings.
    pub fn run(&self, store: &IrStore, root: NodeId) -> RuleReport {
        self.run_with_profile(store, root, &RuleProfile::default())
    }

    /// Runs the rules with a profile and returns findings.
    pub fn run_with_profile(
        &self,
        store: &IrStore,
        root: NodeId,
        profile: &RuleProfile,
    ) -> RuleReport {
        let mut findings = Vec::new();
        let ctx = build_context(store, root, profile);
        for rule in &self.rules {
            if !profile_rule_enabled(profile, rule.id()) {
                continue;
            }
            rule.run(&ctx, &mut findings);
        }

        for finding in &mut findings {
            if let Some(sev) = profile.severity_overrides.get(&finding.rule_id) {
                finding.severity = *sev;
            }
        }

        RuleReport { findings }
    }
}

pub(crate) fn build_context<'a>(
    store: &'a IrStore,
    root: NodeId,
    profile: &RuleProfile,
) -> RuleContext<'a> {
    let mut document = None;
    let mut metadata = None;
    if let Some(IRNode::Document(doc)) = store.get(root) {
        document = Some(doc);
        if let Some(meta_id) = doc.metadata {
            if let Some(IRNode::Metadata(meta)) = store.get(meta_id) {
                metadata = Some(meta);
            }
        }
    }
    RuleContext {
        store,
        root,
        document,
        metadata,
        thresholds: profile.thresholds.clone(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::RuleProfile;
    use docir_core::ir::{Document, DocumentMetadata, IRNode};
    use std::sync::atomic::{AtomicU32, Ordering};
    use std::sync::Arc;

    struct CountingRule {
        id: &'static str,
        calls: Arc<AtomicU32>,
    }

    impl Rule for CountingRule {
        fn id(&self) -> &'static str {
            self.id
        }

        fn name(&self) -> &'static str {
            "counting"
        }

        fn description(&self) -> &'static str {
            "test rule"
        }

        fn category(&self) -> RuleCategory {
            RuleCategory::Security
        }

        fn default_severity(&self) -> Severity {
            Severity::Low
        }

        fn run(&self, _ctx: &RuleContext, findings: &mut Vec<Finding>) {
            self.calls.fetch_add(1, Ordering::Relaxed);
            findings.push(Finding {
                rule_id: self.id.to_string(),
                rule_name: "counting".to_string(),
                category: RuleCategory::Security,
                severity: Severity::Low,
                message: "generated".to_string(),
                context: Vec::new(),
                node_id: None,
                node_type: None,
                location: None,
            });
        }
    }

    fn make_store_with_document(include_metadata: bool) -> (IrStore, NodeId) {
        let mut store = IrStore::new();
        let mut doc = Document::new(docir_core::types::DocumentFormat::WordProcessing);
        if include_metadata {
            let metadata = DocumentMetadata::default();
            let metadata_id = metadata.id;
            doc.metadata = Some(metadata_id);
            store.insert(IRNode::Metadata(metadata));
        }
        let root_id = doc.id;
        store.insert(IRNode::Document(doc));
        (store, root_id)
    }

    #[test]
    fn rule_report_is_empty_reflects_findings_presence() {
        let empty = RuleReport {
            findings: Vec::new(),
        };
        let non_empty = RuleReport {
            findings: vec![Finding {
                rule_id: "X".to_string(),
                rule_name: "x".to_string(),
                category: RuleCategory::Security,
                severity: Severity::Info,
                message: "msg".to_string(),
                context: Vec::new(),
                node_id: None,
                node_type: None,
                location: None,
            }],
        };
        assert!(empty.is_empty());
        assert!(!non_empty.is_empty());
    }

    #[test]
    fn build_context_populates_document_and_metadata_when_available() {
        let (store, root_id) = make_store_with_document(true);
        let ctx = build_context(&store, root_id, &RuleProfile::default());
        assert!(ctx.document.is_some());
        assert!(ctx.metadata.is_some());
    }

    #[test]
    fn build_context_handles_missing_root_document() {
        let store = IrStore::new();
        let root_id = NodeId::new();
        let ctx = build_context(&store, root_id, &RuleProfile::default());
        assert!(ctx.document.is_none());
        assert!(ctx.metadata.is_none());
    }

    #[test]
    fn run_with_profile_applies_disable_and_severity_override() {
        let calls_disabled = Arc::new(AtomicU32::new(0));
        let calls_enabled = Arc::new(AtomicU32::new(0));
        let mut engine = RuleEngine { rules: Vec::new() };
        engine.add_rule(Box::new(CountingRule {
            id: "R-ENABLED",
            calls: calls_enabled.clone(),
        }));
        engine.add_rule(Box::new(CountingRule {
            id: "R-DISABLED",
            calls: calls_disabled.clone(),
        }));

        let (store, root_id) = make_store_with_document(false);
        let mut profile = RuleProfile::default();
        profile.disabled_rules.push("R-DISABLED".to_string());
        profile
            .severity_overrides
            .insert("R-ENABLED".to_string(), Severity::Critical);

        let report = engine.run_with_profile(&store, root_id, &profile);
        assert_eq!(calls_enabled.load(Ordering::Relaxed), 1);
        assert_eq!(calls_disabled.load(Ordering::Relaxed), 0);
        assert_eq!(report.findings.len(), 1);
        assert_eq!(report.findings[0].severity, Severity::Critical);
    }
}
