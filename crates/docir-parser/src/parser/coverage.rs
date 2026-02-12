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
                entries.push(DiagnosticEntry {
                    severity: if status == "complete" {
                        DiagnosticSeverity::Info
                    } else {
                        DiagnosticSeverity::Warning
                    },
                    code: "COVERAGE_PART".to_string(),
                    message: message.clone(),
                    path: Some(path.clone()),
                });
                entries.push(DiagnosticEntry {
                    severity: DiagnosticSeverity::Info,
                    code: "COVERAGE_PART_ROW".to_string(),
                    message,
                    path: Some(path.clone()),
                });
            }
        }
        if !matched {
            missing += 1;
            entries.push(DiagnosticEntry {
                severity: DiagnosticSeverity::Warning,
                code: "COVERAGE_MISSING".to_string(),
                message: format!(
                    "missing part for pattern {} (expected parser={})",
                    spec.pattern, spec.expected_parser
                ),
                path: Some(spec.pattern.to_string()),
            });
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
        entries.push(DiagnosticEntry {
            severity: DiagnosticSeverity::Info,
            code: "CONTENT_TYPE_INVENTORY".to_string(),
            message: format!("content-type: {} => {}", path, ct),
            path: Some(path),
        });
    }
    unknown_ct_paths.sort();
    for path in &unknown_ct_paths {
        entries.push(DiagnosticEntry {
            severity: DiagnosticSeverity::Warning,
            code: "CONTENT_TYPE_UNKNOWN".to_string(),
            message: format!("no content-type entry for path {}", path),
            path: Some(path.clone()),
        });
    }

    entries.push(DiagnosticEntry {
        severity: DiagnosticSeverity::Info,
        code: "COVERAGE_SUMMARY".to_string(),
        message: format!(
            "coverage summary: matched={}, complete={}, pending={}, missing_patterns={}, unknown_content_types={}",
            matched_parts, complete, pending, missing, unknown_ct_paths.len()
        ),
        path: None,
    });
    entries.push(DiagnosticEntry {
        severity: DiagnosticSeverity::Info,
        code: "COVERAGE_COUNTS".to_string(),
        message: format!(
            "counts: complete={}, pending={}, missing_patterns={}, unknown_content_types={}",
            complete,
            pending,
            missing,
            unknown_ct_paths.len()
        ),
        path: None,
    });

    let mut hist: Vec<(String, usize)> = ct_histogram.into_iter().collect();
    hist.sort_by(|a, b| a.0.cmp(&b.0));
    for (ct, count) in hist {
        entries.push(DiagnosticEntry {
            severity: DiagnosticSeverity::Info,
            code: "CONTENT_TYPE_HISTOGRAM".to_string(),
            message: format!("content-type-count: {} => {}", ct, count),
            path: Some(ct),
        });
    }

    entries
}
