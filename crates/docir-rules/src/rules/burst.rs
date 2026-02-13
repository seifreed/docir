use super::{
    count_activex_controls, count_external_links, count_ole_objects,
    threshold_max_activex_controls, threshold_max_external_links, threshold_max_ole_objects,
};
use crate::{Finding, Rule, RuleCategory, RuleContext, RuleThresholds, Severity};

use super::support::add_finding;

pub(super) fn burst_rules() -> Vec<Box<dyn Rule>> {
    vec![
        Box::new(BurstRule::new(&OLE_OBJECT_BURST_RULE)),
        Box::new(BurstRule::new(&ACTIVEX_CONTROL_BURST_RULE)),
        Box::new(BurstRule::new(&EXTERNAL_LINK_BURST_RULE)),
    ]
}

const OLE_OBJECT_BURST_RULE: BurstRuleSpec = BurstRuleSpec {
    id: "SEC-006",
    title: "Excessive OLE objects",
    description: "Detects an unusual number of embedded OLE objects",
    severity: Severity::Medium,
    category: RuleCategory::Security,
    threshold: threshold_max_ole_objects,
    count: count_ole_objects,
    label: "OLE objects",
};

const ACTIVEX_CONTROL_BURST_RULE: BurstRuleSpec = BurstRuleSpec {
    id: "SEC-008",
    title: "Excessive ActiveX controls",
    description: "Detects an unusual number of embedded ActiveX controls",
    severity: Severity::Medium,
    category: RuleCategory::Security,
    threshold: threshold_max_activex_controls,
    count: count_activex_controls,
    label: "ActiveX controls",
};

const EXTERNAL_LINK_BURST_RULE: BurstRuleSpec = BurstRuleSpec {
    id: "SEC-012",
    title: "Excessive external references",
    description: "Detects an unusual number of external references",
    severity: Severity::Medium,
    category: RuleCategory::Security,
    threshold: threshold_max_external_links,
    count: count_external_links,
    label: "External links",
};

struct BurstRuleSpec {
    id: &'static str,
    title: &'static str,
    description: &'static str,
    severity: Severity,
    category: RuleCategory,
    threshold: fn(&RuleThresholds) -> Option<usize>,
    count: fn(&RuleContext) -> usize,
    label: &'static str,
}

struct BurstRule {
    spec: &'static BurstRuleSpec,
}

impl BurstRule {
    fn new(spec: &'static BurstRuleSpec) -> Self {
        Self { spec }
    }
}

impl Rule for BurstRule {
    fn id(&self) -> &'static str {
        self.spec.id
    }

    fn name(&self) -> &'static str {
        self.spec.title
    }

    fn description(&self) -> &'static str {
        self.spec.description
    }

    fn category(&self) -> RuleCategory {
        self.spec.category
    }

    fn default_severity(&self) -> Severity {
        self.spec.severity
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        let Some(threshold) = (self.spec.threshold)(&ctx.thresholds) else {
            return;
        };
        let count = (self.spec.count)(ctx);
        if count > threshold {
            let message = format!("{}: {count} (threshold {threshold})", self.spec.label);
            add_finding(findings, self, message, None, ctx);
        }
    }
}
