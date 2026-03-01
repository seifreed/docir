//! Built-in rules for docir IR analysis.

use crate::{Finding, Rule, RuleCategory, RuleContext, RuleThresholds, Severity};
use docir_core::ir::IRNode;
use docir_core::types::NodeType;

mod burst;
mod structure;
mod support;
use support::{add_finding, is_suspicious_formula, visit_nodes};

pub(crate) fn default_rules() -> Vec<Box<dyn Rule>> {
    let mut rules: Vec<Box<dyn Rule>> = vec![
        Box::new(MacroProjectRule),
        Box::new(MacroAutoExecRule),
        Box::new(SuspiciousVbaCallRule),
        Box::new(XlmMacroRule),
        Box::new(OleObjectRule),
        Box::new(ActiveXControlRule),
        Box::new(DdeFieldRule),
        Box::new(ExternalReferenceRule),
        Box::new(ExternalHyperlinkRule),
        Box::new(SuspiciousFormulaRule),
    ];
    rules.extend(burst::burst_rules());
    rules.extend(structure::structure_rules());
    rules
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
    for_each_external_hyperlink(ctx, |_, _| {
        count += 1;
    });
    count
}

fn add_security_node_findings(
    ctx: &RuleContext,
    findings: &mut Vec<Finding>,
    rule: &dyn Rule,
    ids: &[docir_core::types::NodeId],
    message_for_node: impl Fn(&IRNode) -> Option<String>,
    fallback_message: &'static str,
) {
    for id in ids {
        let node = ctx.store.get(*id);
        let message = node
            .and_then(&message_for_node)
            .unwrap_or_else(|| fallback_message.to_string());
        add_finding(findings, rule, message, node, ctx);
    }
}

fn for_each_external_hyperlink(
    ctx: &RuleContext,
    mut f: impl FnMut(&IRNode, &docir_core::ir::Hyperlink),
) {
    visit_nodes(ctx, |node| {
        if let IRNode::Hyperlink(link) = node {
            if link.is_external {
                f(node, link);
            }
        }
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
        add_security_node_findings(
            ctx,
            findings,
            self,
            &doc.security.ole_objects,
            |node| match node {
                IRNode::OleObject(ole) => Some(format!(
                    "OLE object: {}",
                    ole.prog_id.as_deref().unwrap_or("unknown")
                )),
                _ => None,
            },
            "OLE object detected",
        );
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
        add_security_node_findings(
            ctx,
            findings,
            self,
            &doc.security.activex_controls,
            |node| match node {
                IRNode::ActiveXControl(ctrl) => Some(format!(
                    "ActiveX control detected: {}",
                    ctrl.name.as_deref().unwrap_or("unknown")
                )),
                _ => None,
            },
            "ActiveX control detected",
        );
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
        add_security_node_findings(
            ctx,
            findings,
            self,
            &doc.security.external_refs,
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{RuleEngine, RuleProfile};
    use docir_core::ir::{Cell, CellFormula, Document, FormulaType, Hyperlink, IRNode};
    use docir_core::security::{
        ActiveXControl, DdeField, DdeFieldType, ExternalRefType, ExternalReference, MacroModule,
        MacroModuleType, MacroProject, OleObject, SuspiciousCall, SuspiciousCallCategory, XlmMacro,
    };
    use docir_core::types::DocumentFormat;
    use docir_core::visitor::IrStore;
    use std::collections::HashSet;

    fn build_security_heavy_store() -> (IrStore, docir_core::types::NodeId) {
        let mut store = IrStore::new();
        let mut document = Document::new(DocumentFormat::Spreadsheet);

        let mut macro_project = MacroProject::new();
        macro_project.name = Some("VBAProject".to_string());
        macro_project.has_auto_exec = true;
        macro_project
            .auto_exec_procedures
            .push("AutoOpen".to_string());
        let macro_project_id = macro_project.id;
        store.insert(IRNode::MacroProject(macro_project));
        document.security.macro_project = Some(macro_project_id);

        let mut macro_module = MacroModule::new("Module1", MacroModuleType::Standard);
        macro_module.suspicious_calls.push(SuspiciousCall {
            name: "Shell".to_string(),
            category: SuspiciousCallCategory::ShellExecution,
            line: Some(7),
        });
        store.insert(IRNode::MacroModule(macro_module));

        let mut ole = OleObject::new();
        ole.prog_id = Some("Excel.Sheet.12".to_string());
        let ole_id = ole.id;
        store.insert(IRNode::OleObject(ole));
        document.security.ole_objects.push(ole_id);

        let mut activex = ActiveXControl::new();
        activex.name = Some("CommandButton1".to_string());
        let activex_id = activex.id;
        store.insert(IRNode::ActiveXControl(activex));
        document.security.activex_controls.push(activex_id);

        let ext_ref = ExternalReference::new(ExternalRefType::Hyperlink, "https://evil.test/ref");
        let ext_ref_id = ext_ref.id;
        store.insert(IRNode::ExternalReference(ext_ref));
        document.security.external_refs.push(ext_ref_id);

        document.security.dde_fields.push(DdeField {
            field_type: DdeFieldType::DdeAuto,
            application: "cmd".to_string(),
            topic: Some("/c".to_string()),
            item: Some("calc.exe".to_string()),
            instruction: "DDEAUTO cmd /c calc.exe".to_string(),
            location: None,
        });
        document.security.xlm_macros.push(XlmMacro {
            sheet_name: "Macro1".to_string(),
            sheet_state: docir_core::ir::SheetState::Visible,
            dangerous_functions: Vec::new(),
            macro_cells: Vec::new(),
            has_auto_open: true,
        });

        let link = Hyperlink::new("https://external.example/path", true);
        let link_id = link.id;
        store.insert(IRNode::Hyperlink(link));
        document.content.push(link_id);

        let mut cell = Cell::new("A1", 0, 0);
        cell.formula = Some(CellFormula {
            text: "WEBSERVICE(\"https://exfil.test\")".to_string(),
            formula_type: FormulaType::Normal,
            shared_index: None,
            shared_ref: None,
            is_array: false,
            array_ref: None,
        });
        let cell_id = cell.id;
        store.insert(IRNode::Cell(cell));
        document.content.push(cell_id);

        let root = document.id;
        store.insert(IRNode::Document(document));
        (store, root)
    }

    #[test]
    fn default_rules_expose_stable_metadata() {
        let rules = default_rules();
        assert!(rules.len() >= 13);
        let mut ids = HashSet::new();
        for rule in rules {
            assert!(ids.insert(rule.id().to_string()));
            assert!(!rule.id().is_empty());
            assert!(!rule.name().is_empty());
            assert!(!rule.description().is_empty());
            assert!(matches!(
                rule.category(),
                RuleCategory::Security
                    | RuleCategory::Structure
                    | RuleCategory::Content
                    | RuleCategory::Metadata
            ));
            assert!(matches!(
                rule.default_severity(),
                Severity::Info
                    | Severity::Low
                    | Severity::Medium
                    | Severity::High
                    | Severity::Critical
            ));
        }
    }

    #[test]
    fn default_rules_fire_on_security_heavy_document() {
        let (store, root) = build_security_heavy_store();
        let engine = RuleEngine::with_default_rules();
        let report = engine.run(&store, root);
        let found: HashSet<String> = report.findings.iter().map(|f| f.rule_id.clone()).collect();

        for id in [
            "SEC-001", "SEC-002", "SEC-003", "SEC-004", "SEC-005", "SEC-007", "SEC-009",
            "SEC-010", "SEC-011", "SEC-013",
        ] {
            assert!(found.contains(id), "missing expected finding {id}");
        }
    }

    #[test]
    fn burst_threshold_rules_fire_when_limits_are_low() {
        let (store, root) = build_security_heavy_store();
        let engine = RuleEngine::with_default_rules();
        let mut profile = RuleProfile::default();
        profile.thresholds.max_ole_objects = Some(0);
        profile.thresholds.max_activex_controls = Some(0);
        profile.thresholds.max_external_links = Some(0);

        let report = engine.run_with_profile(&store, root, &profile);
        let found: HashSet<String> = report.findings.iter().map(|f| f.rule_id.clone()).collect();
        assert!(found.contains("SEC-006"));
        assert!(found.contains("SEC-008"));
        assert!(found.contains("SEC-012"));
    }
}
