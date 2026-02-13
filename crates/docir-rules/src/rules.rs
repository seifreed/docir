//! Built-in rules for docir IR analysis.

use crate::{Finding, Rule, RuleCategory, RuleContext, RuleThresholds, Severity};
use docir_core::ir::IRNode;
use docir_core::types::NodeType;

mod support;
use support::{add_finding, is_suspicious_formula, visit_nodes};

pub(crate) fn default_rules() -> Vec<Box<dyn Rule>> {
    vec![
        Box::new(MacroProjectRule),
        Box::new(MacroAutoExecRule),
        Box::new(SuspiciousVbaCallRule),
        Box::new(XlmMacroRule),
        Box::new(OleObjectRule),
        Box::new(BurstRule::new(&OLE_OBJECT_BURST_RULE)),
        Box::new(ActiveXControlRule),
        Box::new(BurstRule::new(&ACTIVEX_CONTROL_BURST_RULE)),
        Box::new(DdeFieldRule),
        Box::new(ExternalReferenceRule),
        Box::new(ExternalHyperlinkRule),
        Box::new(BurstRule::new(&EXTERNAL_LINK_BURST_RULE)),
        Box::new(HiddenWorksheetRule),
        Box::new(HiddenSlideRule),
        Box::new(SuspiciousFormulaRule),
    ]
}

fn threshold_max_ole_objects(thresholds: &RuleThresholds) -> Option<usize> {
    thresholds.max_ole_objects
}

fn threshold_max_activex_controls(thresholds: &RuleThresholds) -> Option<usize> {
    thresholds.max_activex_controls
}

fn threshold_max_external_links(thresholds: &RuleThresholds) -> Option<usize> {
    thresholds.max_external_links
}

fn count_ole_objects(ctx: &RuleContext) -> usize {
    ctx.document
        .map(|doc| doc.security.ole_objects.len())
        .unwrap_or(0)
}

fn count_activex_controls(ctx: &RuleContext) -> usize {
    ctx.document
        .map(|doc| doc.security.activex_controls.len())
        .unwrap_or(0)
}

fn count_external_links(ctx: &RuleContext) -> usize {
    let mut count = 0usize;
    visit_nodes(ctx, |node| {
        if let IRNode::Hyperlink(link) = node {
            if link.is_external {
                count += 1;
            }
        }
    });
    count
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

/// Rule: Macro project present.
struct MacroProjectRule;

impl Rule for MacroProjectRule {
    fn id(&self) -> &'static str {
        "SEC-001"
    }

    fn name(&self) -> &'static str {
        "Macro project present"
    }

    fn description(&self) -> &'static str {
        "Detects presence of VBA macro project"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::Critical
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        let Some(doc) = ctx.document else {
            return;
        };
        if let Some(macro_id) = doc.security.macro_project {
            let node = ctx.store.get(macro_id);
            add_finding(
                findings,
                self,
                "VBA macro project detected".to_string(),
                node,
                ctx,
            );
        }
    }
}

/// Rule: Macro auto-exec procedures.
struct MacroAutoExecRule;

impl Rule for MacroAutoExecRule {
    fn id(&self) -> &'static str {
        "SEC-002"
    }

    fn name(&self) -> &'static str {
        "Macro auto-exec detected"
    }

    fn description(&self) -> &'static str {
        "Detects auto-exec procedures in VBA macro project"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::Critical
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        let Some(doc) = ctx.document else {
            return;
        };
        let Some(macro_id) = doc.security.macro_project else {
            return;
        };
        let Some(IRNode::MacroProject(project)) = ctx.store.get(macro_id) else {
            return;
        };

        if project.has_auto_exec {
            add_finding(
                findings,
                self,
                format!("Auto-exec macros: {:?}", project.auto_exec_procedures),
                Some(&IRNode::MacroProject(project.clone())),
                ctx,
            );
        }
    }
}

/// Rule: Suspicious VBA API calls.
struct SuspiciousVbaCallRule;

impl Rule for SuspiciousVbaCallRule {
    fn id(&self) -> &'static str {
        "SEC-003"
    }

    fn name(&self) -> &'static str {
        "Suspicious VBA API"
    }

    fn description(&self) -> &'static str {
        "Detects potentially dangerous VBA API calls"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::High
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        for module_id in ctx.store.iter_ids_by_type(NodeType::MacroModule) {
            let Some(IRNode::MacroModule(module)) = ctx.store.get(module_id) else {
                continue;
            };
            for call in &module.suspicious_calls {
                add_finding(
                    findings,
                    self,
                    format!("Suspicious VBA call: {}", call.name),
                    Some(&IRNode::MacroModule(module.clone())),
                    ctx,
                );
            }
        }
    }
}

/// Rule: XLM macros present.
struct XlmMacroRule;

impl Rule for XlmMacroRule {
    fn id(&self) -> &'static str {
        "SEC-004"
    }

    fn name(&self) -> &'static str {
        "XLM macros present"
    }

    fn description(&self) -> &'static str {
        "Detects Excel 4.0 XLM macros"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::High
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        let Some(doc) = ctx.document else {
            return;
        };
        if !doc.security.xlm_macros.is_empty() {
            add_finding(
                findings,
                self,
                format!("XLM macros detected: {}", doc.security.xlm_macros.len()),
                None,
                ctx,
            );
        }
    }
}

/// Rule: OLE object present.
struct OleObjectRule;

impl Rule for OleObjectRule {
    fn id(&self) -> &'static str {
        "SEC-005"
    }

    fn name(&self) -> &'static str {
        "OLE object present"
    }

    fn description(&self) -> &'static str {
        "Detects embedded OLE objects"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::High
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        let Some(doc) = ctx.document else {
            return;
        };
        if doc.security.ole_objects.is_empty() {
            return;
        }
        for ole_id in &doc.security.ole_objects {
            let node = ctx.store.get(*ole_id);
            if let Some(IRNode::OleObject(ole)) = node {
                add_finding(
                    findings,
                    self,
                    format!(
                        "OLE object: {}",
                        ole.prog_id.as_deref().unwrap_or("unknown")
                    ),
                    node,
                    ctx,
                );
            } else {
                add_finding(findings, self, "OLE object detected".to_string(), node, ctx);
            }
        }
    }
}

/// Rule: ActiveX control present.
struct ActiveXControlRule;

impl Rule for ActiveXControlRule {
    fn id(&self) -> &'static str {
        "SEC-007"
    }

    fn name(&self) -> &'static str {
        "ActiveX control present"
    }

    fn description(&self) -> &'static str {
        "Detects embedded ActiveX controls"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::High
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        let Some(doc) = ctx.document else {
            return;
        };
        if doc.security.activex_controls.is_empty() {
            return;
        }
        for id in &doc.security.activex_controls {
            let node = ctx.store.get(*id);
            if let Some(IRNode::ActiveXControl(ctrl)) = node {
                add_finding(
                    findings,
                    self,
                    format!(
                        "ActiveX control detected: {}",
                        ctrl.name.as_deref().unwrap_or("unknown")
                    ),
                    node,
                    ctx,
                );
            } else {
                add_finding(
                    findings,
                    self,
                    "ActiveX control detected".to_string(),
                    node,
                    ctx,
                );
            }
        }
    }
}

/// Rule: DDE field present.
struct DdeFieldRule;

impl Rule for DdeFieldRule {
    fn id(&self) -> &'static str {
        "SEC-009"
    }

    fn name(&self) -> &'static str {
        "DDE field present"
    }

    fn description(&self) -> &'static str {
        "Detects DDE fields"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::High
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        let Some(doc) = ctx.document else {
            return;
        };
        if doc.security.dde_fields.is_empty() {
            return;
        }
        for field in &doc.security.dde_fields {
            add_finding(
                findings,
                self,
                format!("DDE field: {}", field.instruction),
                None,
                ctx,
            );
        }
    }
}

/// Rule: External reference present.
struct ExternalReferenceRule;

impl Rule for ExternalReferenceRule {
    fn id(&self) -> &'static str {
        "SEC-010"
    }

    fn name(&self) -> &'static str {
        "External references present"
    }

    fn description(&self) -> &'static str {
        "Detects external references"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::Low
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        let Some(doc) = ctx.document else {
            return;
        };
        if doc.security.external_refs.is_empty() {
            return;
        }
        for ref_id in &doc.security.external_refs {
            let node = ctx.store.get(*ref_id);
            if let Some(IRNode::ExternalReference(ext)) = node {
                add_finding(
                    findings,
                    self,
                    format!("External reference: {}", ext.target),
                    node,
                    ctx,
                );
            } else {
                add_finding(
                    findings,
                    self,
                    "External reference detected".to_string(),
                    node,
                    ctx,
                );
            }
        }
    }
}

/// Rule: External hyperlinks in document content.
struct ExternalHyperlinkRule;

impl Rule for ExternalHyperlinkRule {
    fn id(&self) -> &'static str {
        "SEC-011"
    }

    fn name(&self) -> &'static str {
        "External hyperlinks"
    }

    fn description(&self) -> &'static str {
        "Detects external hyperlinks in document content"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::Low
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        visit_nodes(ctx, |node| {
            if let IRNode::Hyperlink(link) = node {
                if link.is_external {
                    add_finding(
                        findings,
                        self,
                        format!("External hyperlink: {}", link.target),
                        Some(node),
                        ctx,
                    );
                }
            }
        });
    }
}

/// Rule: Hidden worksheets.
struct HiddenWorksheetRule;

impl Rule for HiddenWorksheetRule {
    fn id(&self) -> &'static str {
        "STR-001"
    }

    fn name(&self) -> &'static str {
        "Hidden worksheets"
    }

    fn description(&self) -> &'static str {
        "Detects hidden worksheets"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Structure
    }

    fn default_severity(&self) -> Severity {
        Severity::Medium
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        visit_nodes(ctx, |node| {
            if let IRNode::Worksheet(sheet) = node {
                if sheet.state != docir_core::ir::SheetState::Visible {
                    add_finding(
                        findings,
                        self,
                        format!("Hidden worksheet: {}", sheet.name),
                        Some(node),
                        ctx,
                    );
                }
            }
        });
    }
}

/// Rule: Hidden slides.
struct HiddenSlideRule;

impl Rule for HiddenSlideRule {
    fn id(&self) -> &'static str {
        "STR-002"
    }

    fn name(&self) -> &'static str {
        "Hidden slides"
    }

    fn description(&self) -> &'static str {
        "Detects hidden slides"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Structure
    }

    fn default_severity(&self) -> Severity {
        Severity::Low
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        visit_nodes(ctx, |node| {
            if let IRNode::Slide(slide) = node {
                if slide.hidden {
                    add_finding(
                        findings,
                        self,
                        format!("Hidden slide: {}", slide.number),
                        Some(node),
                        ctx,
                    );
                }
            }
        });
    }
}

/// Rule: Suspicious formulas.
struct SuspiciousFormulaRule;

impl Rule for SuspiciousFormulaRule {
    fn id(&self) -> &'static str {
        "SEC-013"
    }

    fn name(&self) -> &'static str {
        "Suspicious formulas"
    }

    fn description(&self) -> &'static str {
        "Detects formulas that invoke external or shell-like functions"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::High
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        visit_nodes(ctx, |node| {
            if let IRNode::Cell(cell) = node {
                if let Some(formula) = &cell.formula {
                    if is_suspicious_formula(&formula.text) {
                        add_finding(
                            findings,
                            self,
                            format!("Suspicious formula in {}: {}", cell.reference, formula.text),
                            Some(node),
                            ctx,
                        );
                    }
                }
            }
        });
    }
}
