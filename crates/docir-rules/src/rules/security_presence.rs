use super::{
    for_each_external_hyperlink, run_activex_control_rule, run_ole_object_rule,
    run_security_node_rule,
};
use crate::{Finding, Rule, RuleCategory, RuleContext, Severity};
use docir_core::ir::IRNode;
use docir_core::types::NodeType;

use super::support::add_finding;

pub(super) fn security_presence_rules() -> Vec<Box<dyn Rule>> {
    vec![
        Box::new(MacroProjectRule),
        Box::new(MacroAutoExecRule),
        Box::new(SuspiciousVbaCallRule),
        Box::new(XlmMacroRule),
        Box::new(OleObjectRule),
        Box::new(ActiveXControlRule),
        Box::new(DdeFieldRule),
        Box::new(ExternalReferenceRule),
        Box::new(ExternalHyperlinkRule),
    ]
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
        if let Some(macro_id) = doc.security.macro_project_id() {
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
        let Some(macro_id) = doc.security.macro_project_id() else {
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
        if doc.security.has_xlm_macros() {
            add_finding(
                findings,
                self,
                format!("XLM macros detected: {}", doc.security.xlm_macro_count()),
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
        run_ole_object_rule(ctx, findings, self);
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
        run_activex_control_rule(ctx, findings, self);
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
        if !doc.security.has_dde_fields() {
            return;
        }
        for field in doc.security.dde_fields() {
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
        run_security_node_rule(
            ctx,
            findings,
            self,
            |doc| doc.security.has_external_references(),
            |doc| doc.security.external_ref_ids(),
            |node| match node {
                IRNode::ExternalReference(ext) => {
                    Some(format!("External reference: {}", ext.target))
                }
                _ => None,
            },
            "External reference detected",
        );
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
        for_each_external_hyperlink(ctx, |node, link| {
            add_finding(
                findings,
                self,
                format!("External hyperlink: {}", link.target),
                Some(node),
                ctx,
            );
        });
    }
}
