use crate::{Finding, Rule, RuleCategory, RuleContext, Severity};
use docir_core::ir::IRNode;

use super::support::{add_finding, is_suspicious_formula, visit_nodes};

pub(super) fn security_formula_rules() -> Vec<Box<dyn Rule>> {
    vec![Box::new(SuspiciousFormulaRule)]
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
