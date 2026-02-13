use crate::{Finding, Rule, RuleCategory, RuleContext, Severity};
use docir_core::ir::IRNode;

use super::support::{add_finding, visit_nodes};

pub(super) fn structure_rules() -> Vec<Box<dyn Rule>> {
    vec![Box::new(HiddenWorksheetRule), Box::new(HiddenSlideRule)]
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
