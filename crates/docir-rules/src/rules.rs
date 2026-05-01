//! Built-in rules for docir IR analysis.

use crate::{Finding, Rule, RuleContext, RuleThresholds};
use docir_core::ir::IRNode;

mod burst;
mod security_formula;
mod security_presence;
mod structure;
mod support;
use security_formula::security_formula_rules;
use security_presence::security_presence_rules;
use support::{add_finding, visit_nodes};

pub(crate) fn default_rules() -> Vec<Box<dyn Rule>> {
    let mut rules: Vec<Box<dyn Rule>> = Vec::new();
    rules.extend(security_presence_rules());
    rules.extend(security_formula_rules());
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
        .map(|doc| doc.security.ole_object_count())
        .unwrap_or(0)
}

fn count_activex_controls(ctx: &RuleContext) -> usize {
    ctx.document
        .map(|doc| doc.security.activex_control_count())
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

pub(super) fn run_security_node_rule(
    ctx: &RuleContext,
    findings: &mut Vec<Finding>,
    rule: &dyn Rule,
    enabled: impl FnOnce(&docir_core::ir::Document) -> bool,
    ids: impl FnOnce(&docir_core::ir::Document) -> &[docir_core::types::NodeId],
    message_for_node: impl Fn(&IRNode) -> Option<String>,
    fallback_message: &'static str,
) {
    let Some(doc) = ctx.document else {
        return;
    };
    if !enabled(doc) {
        return;
    }
    add_security_node_findings(
        ctx,
        findings,
        rule,
        ids(doc),
        message_for_node,
        fallback_message,
    );
}

pub(super) fn run_ole_object_rule(ctx: &RuleContext, findings: &mut Vec<Finding>, rule: &dyn Rule) {
    run_security_node_rule(
        ctx,
        findings,
        rule,
        |doc| doc.security.has_ole_objects(),
        |doc| doc.security.ole_object_ids(),
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

pub(super) fn run_activex_control_rule(
    ctx: &RuleContext,
    findings: &mut Vec<Finding>,
    rule: &dyn Rule,
) {
    run_security_node_rule(
        ctx,
        findings,
        rule,
        |doc| doc.security.has_activex_controls(),
        |doc| doc.security.activex_control_ids(),
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

pub(super) fn for_each_external_hyperlink(
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{RuleCategory, Severity};
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
            "SEC-001", "SEC-002", "SEC-003", "SEC-004", "SEC-005", "SEC-007", "SEC-009", "SEC-010",
            "SEC-011", "SEC-013",
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
