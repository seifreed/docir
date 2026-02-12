use crate::diagnostics::push_entry;
use crate::ooxml::content_types::ContentTypes;
use crate::ooxml::part_registry;
use docir_core::ir::{DiagnosticEntry, DiagnosticSeverity, IRNode, IrNode as IrNodeTrait};
use docir_core::types::DocumentFormat;
use docir_core::visitor::IrStore;
use std::collections::{HashMap, HashSet};

pub(super) fn collect_seen_paths(store: &IrStore) -> HashSet<String> {
    let mut seen = HashSet::new();
    for node in store.values() {
        if let Some(span) = node.source_span() {
            seen.insert(span.file_path.clone());
        }
        if let IRNode::MediaAsset(media) = node {
            seen.insert(media.path.clone());
        }
        if let IRNode::CustomXmlPart(part) = node {
            seen.insert(part.path.clone());
        }
        if let IRNode::ExtensionPart(part) = node {
            seen.insert(part.path.clone());
        }
        if let IRNode::RelationshipGraph(graph) = node {
            seen.insert(graph.source.clone());
        }
    }
    seen
}

pub(super) fn build_coverage_diagnostics(
    format: DocumentFormat,
    content_types: &ContentTypes,
    all_paths: Vec<String>,
    seen_paths: &HashSet<String>,
) -> Vec<DiagnosticEntry> {
    let all_paths: Vec<String> = all_paths
        .into_iter()
        .filter(|p| !p.ends_with('/') && !p.starts_with("[trash]/"))
        .collect();
    let registry = part_registry::registry_for(format);
    let mut entries = Vec::new();

    let mut matched_parts = 0usize;
    let mut complete = 0usize;
    let mut pending = 0usize;
    let mut missing = 0usize;
    let mut by_status: HashMap<&'static str, usize> = HashMap::new();

    for spec in registry {
        let mut matched = false;
        for path in &all_paths {
            let ct = content_types.get_content_type(path);
            if spec.matches(path, ct) {
                matched = true;
                matched_parts += 1;
                let status = if seen_paths.contains(path) {
                    complete += 1;
                    "complete"
                } else {
                    pending += 1;
                    "pending"
                };
                *by_status.entry(status).or_insert(0) += 1;
                let ct_str = ct.unwrap_or("unknown");
                let message = format!(
                    "{} part: {} (content-type={}, parser={})",
                    status, path, ct_str, spec.expected_parser
                );
                let severity = if status == "complete" {
                    DiagnosticSeverity::Info
                } else {
                    DiagnosticSeverity::Warning
                };
                push_entry(
                    &mut entries,
                    severity,
                    "COVERAGE_PART",
                    message.clone(),
                    Some(path),
                );
                push_entry(
                    &mut entries,
                    DiagnosticSeverity::Info,
                    "COVERAGE_PART_ROW",
                    message,
                    Some(path),
                );
            }
        }
        if !matched {
            missing += 1;
            push_entry(
                &mut entries,
                DiagnosticSeverity::Warning,
                "COVERAGE_MISSING",
                format!(
                    "missing part for pattern {} (expected parser={})",
                    spec.pattern, spec.expected_parser
                ),
                Some(spec.pattern),
            );
        }
    }

    let mut inventory: Vec<(String, String)> = Vec::new();
    let mut unknown_ct_paths: Vec<String> = Vec::new();
    let mut ct_histogram: HashMap<String, usize> = HashMap::new();
    for path in &all_paths {
        if let Some(ct) = content_types.get_content_type(path) {
            inventory.push((path.clone(), ct.to_string()));
            *ct_histogram.entry(ct.to_string()).or_insert(0) += 1;
        } else {
            unknown_ct_paths.push(path.clone());
        }
    }
    inventory.sort_by(|a, b| a.0.cmp(&b.0));
    for (path, ct) in inventory {
        push_entry(
            &mut entries,
            DiagnosticSeverity::Info,
            "CONTENT_TYPE_INVENTORY",
            format!("content-type: {} => {}", path, ct),
            Some(&path),
        );
    }
    unknown_ct_paths.sort();
    for path in &unknown_ct_paths {
        push_entry(
            &mut entries,
            DiagnosticSeverity::Warning,
            "CONTENT_TYPE_UNKNOWN",
            format!("no content-type entry for path {}", path),
            Some(path),
        );
    }

    push_entry(
        &mut entries,
        DiagnosticSeverity::Info,
        "COVERAGE_SUMMARY",
        format!(
            "coverage summary: matched={}, complete={}, pending={}, missing_patterns={}, unknown_content_types={}",
            matched_parts, complete, pending, missing, unknown_ct_paths.len()
        ),
        None,
    );
    push_entry(
        &mut entries,
        DiagnosticSeverity::Info,
        "COVERAGE_COUNTS",
        format!(
            "counts: complete={}, pending={}, missing_patterns={}, unknown_content_types={}",
            complete,
            pending,
            missing,
            unknown_ct_paths.len()
        ),
        None,
    );

    let mut hist: Vec<(String, usize)> = ct_histogram.into_iter().collect();
    hist.sort_by(|a, b| a.0.cmp(&b.0));
    for (ct, count) in hist {
        push_entry(
            &mut entries,
            DiagnosticSeverity::Info,
            "CONTENT_TYPE_HISTOGRAM",
            format!("content-type-count: {} => {}", ct, count),
            Some(&ct),
        );
    }

    entries
}
