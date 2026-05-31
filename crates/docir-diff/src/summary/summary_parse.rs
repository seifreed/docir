use docir_core::ir::IRNode;
use docir_core::visitor::IrStore;

use super::summary_primary::summarize_primary;
use super::summary_secondary::summarize_secondary;
use super::{presentation, spreadsheet};

pub(crate) fn summarize(node: &IRNode, store: &IrStore) -> String {
    if let Some(summary) = spreadsheet::summarize(node, store) {
        return summary;
    }
    if let Some(summary) = presentation::summarize(node, store) {
        return summary;
    }
    summarize_with_fallback(node, store)
}

fn summarize_with_fallback(node: &IRNode, store: &IrStore) -> String {
    summarize_primary(node, store).unwrap_or_else(|| summarize_secondary(node))
}
