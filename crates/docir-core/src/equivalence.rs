//! IR equivalence summaries for cross-format parity checks.

use crate::ir::{IRNode, IrNode as IrNodeTrait};
use crate::types::NodeType;
use crate::visitor::IrStore;
use std::collections::{BTreeSet, HashMap};

/// Lightweight IR summary for parity checks.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct IrSummary {
    pub node_counts: HashMap<NodeType, usize>,
    pub run_texts: BTreeSet<String>,
}

impl IrSummary {
    /// Builds a summary from an IR store.
    pub fn from_store(store: &IrStore) -> Self {
        let mut summary = IrSummary::default();
        for node in store.values() {
            let node_type = node.node_type();
            *summary.node_counts.entry(node_type).or_insert(0) += 1;

            if let IRNode::Run(run) = node {
                if !run.text.is_empty() {
                    summary.run_texts.insert(run.text.clone());
                }
            }
        }
        summary
    }

    /// Gets the count for a node type.
    pub fn count(&self, node_type: NodeType) -> usize {
        *self.node_counts.get(&node_type).unwrap_or(&0)
    }
}
