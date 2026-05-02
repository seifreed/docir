use crate::{Finding, Rule, RuleContext};
use docir_core::ir::{IRNode, IrNode as IrNodeTrait};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;
use std::collections::HashSet;

pub(super) fn add_finding(
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

pub(super) fn visit_nodes(ctx: &RuleContext, mut visitor: impl FnMut(&IRNode)) {
    let mut visited = HashSet::new();
    for node in iter_nodes(ctx.store, ctx.root, &mut visited) {
        visitor(node);
    }
}

pub(super) fn is_suspicious_formula(text: &str) -> bool {
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
    let code_only = strip_string_literals(&upper);
    tokens.iter().any(|t| {
        for pos in code_only.match_indices(t) {
            let (idx, _) = pos;
            let before_ok = idx == 0
                || !code_only
                    .as_bytes()
                    .get(idx - 1)
                    .is_some_and(|b| b.is_ascii_alphanumeric() || *b == b'_');
            if before_ok {
                return true;
            }
        }
        false
    })
}

fn strip_string_literals(input: &str) -> String {
    let mut result = String::with_capacity(input.len());
    let mut in_string = false;
    let mut chars = input.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '"' {
            if in_string {
                if chars.peek() == Some(&'"') {
                    chars.next();
                    result.push(' ');
                } else {
                    in_string = false;
                }
            } else {
                in_string = true;
            }
            result.push(' ');
        } else if in_string {
            result.push(' ');
        } else {
            result.push(ch);
        }
    }
    result
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
