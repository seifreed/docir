//! Structured query helpers for the IR.

use crate::ir::{IRNode, IrNode as IrNodeTrait};
use crate::types::{DocumentFormat, NodeId, NodeType};
use crate::visitor::IrStore;
use std::collections::HashSet;

/// Query predicate for IR search.
#[derive(Debug, Clone)]
pub struct Query {
    pub node_types: Vec<NodeType>,
    pub text_contains: Option<String>,
    pub format: Option<DocumentFormat>,
    pub has_external_refs: Option<bool>,
    pub has_macros: Option<bool>,
}

impl Query {
    /// Create an empty query (matches all nodes).
    pub fn new() -> Self {
        Self {
            node_types: Vec::new(),
            text_contains: None,
            format: None,
            has_external_refs: None,
            has_macros: None,
        }
    }

    /// Execute query against an IR tree.
    pub fn execute(&self, store: &IrStore, root: NodeId) -> Vec<NodeId> {
        let mut results = Vec::new();
        let mut visited = HashSet::new();
        let mut stack = vec![root];

        while let Some(id) = stack.pop() {
            if !visited.insert(id) {
                continue;
            }
            let Some(node) = store.get(id) else {
                continue;
            };

            if self.matches_node(node, store) {
                results.push(id);
            }

            for child in node.children().into_iter().rev() {
                stack.push(child);
            }
        }

        results
    }

    fn matches_node(&self, node: &IRNode, store: &IrStore) -> bool {
        if !self.node_types.is_empty() && !self.node_types.contains(&node.node_type()) {
            return false;
        }

        if let Some(format) = self.format {
            let matches_format = match node {
                IRNode::Document(doc) => doc.format == format,
                _ => true,
            };
            if !matches_format {
                return false;
            }
        }

        if let Some(expected) = self.has_external_refs {
            let has_refs = match node {
                IRNode::Document(doc) => doc.security.has_external_references(),
                _ => true,
            };
            if has_refs != expected {
                return false;
            }
        }

        if let Some(expected) = self.has_macros {
            let has_macros = match node {
                IRNode::Document(doc) => doc.security.has_macros(),
                _ => true,
            };
            if has_macros != expected {
                return false;
            }
        }

        if let Some(text) = &self.text_contains {
            let hay = node_text(node, store).unwrap_or_default();
            if !hay.to_lowercase().contains(&text.to_lowercase()) {
                return false;
            }
        }

        true
    }
}

fn node_text(node: &IRNode, store: &IrStore) -> Option<String> {
    match node {
        IRNode::Run(run) => Some(run.text.clone()),
        IRNode::Paragraph(para) => {
            let mut out = String::new();
            for run_id in &para.runs {
                if let Some(IRNode::Run(run)) = store.get(*run_id) {
                    out.push_str(&run.text);
                }
            }
            Some(out)
        }
        IRNode::Shape(shape) => shape.text.as_ref().map(|t| {
            let mut out = String::new();
            for para in &t.paragraphs {
                for run in &para.runs {
                    out.push_str(&run.text);
                }
                out.push('\n');
            }
            out
        }),
        IRNode::Cell(cell) => {
            if let Some(formula) = &cell.formula {
                Some(formula.text.clone())
            } else {
                Some(format!("{:?}", cell.value))
            }
        }
        IRNode::Hyperlink(link) => Some(link.target.clone()),
        _ => None,
    }
}

impl Default for Query {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::{Document, Paragraph, Run};
    use crate::types::{DocumentFormat, NodeType};
    use crate::visitor::IrStore;

    #[test]
    fn query_text_contains_matches_runs() {
        let mut store = IrStore::new();
        let mut doc = Document::new(DocumentFormat::WordProcessing);

        let mut para = Paragraph::new();
        let run = Run::new("hello world");
        let run_id = run.id;
        para.runs.push(run_id);
        let para_id = para.id;

        store.insert(IRNode::Run(run));
        store.insert(IRNode::Paragraph(para));
        doc.content.push(para_id);
        let doc_id = doc.id;
        store.insert(IRNode::Document(doc));

        let mut query = Query::new();
        query.text_contains = Some("world".to_string());
        query.node_types.push(NodeType::Paragraph);

        let results = query.execute(&store, doc_id);
        assert_eq!(results.len(), 1);
    }
}
