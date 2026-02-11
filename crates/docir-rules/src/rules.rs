//! Built-in rules for docir IR analysis.

use crate::{Finding, Rule, RuleCategory, RuleContext, Severity};
use docir_core::ir::{IRNode, IrNode as IrNodeTrait};
use docir_core::types::{NodeId, NodeType};
use docir_core::visitor::IrStore;
use std::collections::HashSet;

pub(crate) fn default_rules() -> Vec<Box<dyn Rule>> {
    vec![
        Box::new(MacroProjectRule),
        Box::new(MacroAutoExecRule),
        Box::new(SuspiciousVbaCallRule),
        Box::new(XlmMacroRule),
        Box::new(OleObjectRule),
        Box::new(OleObjectBurstRule),
        Box::new(ActiveXControlRule),
        Box::new(ActiveXControlBurstRule),
        Box::new(DdeFieldRule),
        Box::new(ExternalReferenceRule),
        Box::new(ExternalHyperlinkRule),
        Box::new(ExternalLinkBurstRule),
        Box::new(HiddenWorksheetRule),
        Box::new(HiddenSlideRule),
        Box::new(SuspiciousFormulaRule),
    ]
}

fn add_finding(
    findings: &mut Vec<Finding>,
    rule: &dyn Rule,
    message: String,
    node: Option<&IRNode>,
    ctx: &RuleContext,
) {
    let (node_id, node_type, location) = node
        .map(|n| {
            let span = n.source_span();
            (
                Some(n.node_id()),
                Some(n.node_type()),
                span.map(|s| s.file_path.clone()),
            )
        })
        .unwrap_or((None, None, None));

    let mut context = Vec::new();
    if let Some(doc) = ctx.document {
        context.push(format!("format={:?}", doc.format));
    }
    if let Some(meta) = ctx.metadata {
        if let Some(title) = &meta.title {
            context.push(format!("title={title}"));
        }
        if let Some(author) = &meta.creator {
            context.push(format!("author={author}"));
        }
    }

    findings.push(Finding {
        rule_id: rule.id().to_string(),
        rule_name: rule.name().to_string(),
        category: rule.category(),
        severity: rule.default_severity(),
        message,
        context,
        node_id,
        node_type,
        location,
    });
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

/// Rule: Excessive OLE objects.
struct OleObjectBurstRule;

impl Rule for OleObjectBurstRule {
    fn id(&self) -> &'static str {
        "SEC-006"
    }

    fn name(&self) -> &'static str {
        "Excessive OLE objects"
    }

    fn description(&self) -> &'static str {
        "Detects an unusual number of embedded OLE objects"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::Medium
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        let Some(doc) = ctx.document else {
            return;
        };
        let Some(threshold) = ctx.thresholds.max_ole_objects else {
            return;
        };
        let count = doc.security.ole_objects.len();
        if count > threshold {
            let message = format!("OLE objects: {count} (threshold {threshold})");
            add_finding(findings, self, message, None, ctx);
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

/// Rule: Excessive ActiveX controls.
struct ActiveXControlBurstRule;

impl Rule for ActiveXControlBurstRule {
    fn id(&self) -> &'static str {
        "SEC-008"
    }

    fn name(&self) -> &'static str {
        "Excessive ActiveX controls"
    }

    fn description(&self) -> &'static str {
        "Detects an unusual number of embedded ActiveX controls"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::Medium
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        let Some(doc) = ctx.document else {
            return;
        };
        let Some(threshold) = ctx.thresholds.max_activex_controls else {
            return;
        };
        let count = doc.security.activex_controls.len();
        if count > threshold {
            let message = format!("ActiveX controls: {count} (threshold {threshold})");
            add_finding(findings, self, message, None, ctx);
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
        let mut visited = HashSet::new();
        for node in iter_nodes(ctx.store, ctx.root, &mut visited) {
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
        }
    }
}

/// Rule: Excessive external references.
struct ExternalLinkBurstRule;

impl Rule for ExternalLinkBurstRule {
    fn id(&self) -> &'static str {
        "SEC-012"
    }

    fn name(&self) -> &'static str {
        "Excessive external references"
    }

    fn description(&self) -> &'static str {
        "Detects an unusual number of external references"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Security
    }

    fn default_severity(&self) -> Severity {
        Severity::Medium
    }

    fn run(&self, ctx: &RuleContext, findings: &mut Vec<Finding>) {
        let Some(threshold) = ctx.thresholds.max_external_links else {
            return;
        };
        let mut count = 0usize;
        let mut visited = HashSet::new();
        for node in iter_nodes(ctx.store, ctx.root, &mut visited) {
            if let IRNode::Hyperlink(link) = node {
                if link.is_external {
                    count += 1;
                }
            }
        }
        if count > threshold {
            let message = format!("External links: {count} (threshold {threshold})");
            add_finding(findings, self, message, None, ctx);
        }
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
        let mut visited = HashSet::new();
        for node in iter_nodes(ctx.store, ctx.root, &mut visited) {
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
        }
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
        let mut visited = HashSet::new();
        for node in iter_nodes(ctx.store, ctx.root, &mut visited) {
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
        }
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
        let mut visited = HashSet::new();
        for node in iter_nodes(ctx.store, ctx.root, &mut visited) {
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
        }
    }
}

fn is_suspicious_formula(text: &str) -> bool {
    let upper = text.to_ascii_uppercase();
    let tokens = [
        "WEBSERVICE(",
        "HYPERLINK(",
        "URL(",
        "EXEC(",
        "CALL(",
        "SHELL(",
        "DDE(",
        "DDEAUTO(",
    ];
    tokens.iter().any(|t| upper.contains(t))
}

fn iter_nodes<'a>(
    store: &'a IrStore,
    root: NodeId,
    visited: &'a mut HashSet<NodeId>,
) -> Vec<&'a IRNode> {
    let mut out = Vec::new();
    let mut stack = vec![root];

    while let Some(id) = stack.pop() {
        if !visited.insert(id) {
            continue;
        }
        let Some(node) = store.get(id) else {
            continue;
        };
        out.push(node);
        for child in node.children().into_iter().rev() {
            stack.push(child);
        }
    }

    out
}
