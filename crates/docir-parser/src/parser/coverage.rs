use crate::diagnostics::push_entry;
use crate::ooxml::content_types::ContentTypes;
use crate::ooxml::part_registry;
use docir_core::ir::{DiagnosticEntry, DiagnosticSeverity, IRNode, IrNode as IrNodeTrait};
use docir_core::types::DocumentFormat;
use docir_core::visitor::IrStore;
use std::collections::{HashMap, HashSet};

type ContentTypeInventory = (Vec<(String, String)>, Vec<String>, HashMap<String, usize>);

#[derive(Default)]
struct CoverageCounts {
    matched_parts: usize,
    complete: usize,
    pending: usize,
    missing: usize,
}

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
    let all_paths = filter_paths(all_paths);
    let registry = part_registry::registry_for(format);
    let mut entries = Vec::new();
    let counts = emit_registry_diagnostics(
        &mut entries,
        &registry,
        content_types,
        &all_paths,
        seen_paths,
    );
    let (inventory, unknown_ct_paths, ct_histogram) =
        collect_content_type_data(content_types, &all_paths);
    emit_content_type_diagnostics(&mut entries, inventory, &unknown_ct_paths);
    emit_summary_diagnostics(&mut entries, &counts, unknown_ct_paths.len());
    emit_histogram_diagnostics(&mut entries, ct_histogram);

    entries
}

fn filter_paths(all_paths: Vec<String>) -> Vec<String> {
    all_paths
        .into_iter()
        .filter(|p| !p.ends_with('/') && !p.starts_with("[trash]/"))
        .collect()
}

fn emit_registry_diagnostics(
    entries: &mut Vec<DiagnosticEntry>,
    registry: &[part_registry::PartSpec],
    content_types: &ContentTypes,
    all_paths: &[String],
    seen_paths: &HashSet<String>,
) -> CoverageCounts {
    let mut counts = CoverageCounts::default();

    for spec in registry {
        let mut matched = false;
        for path in all_paths {
            let ct = content_types.get_content_type(path);
            if spec.matches(path, ct) {
                matched = true;
                counts.matched_parts += 1;
                let (status, severity) = if seen_paths.contains(path) {
                    counts.complete += 1;
                    ("complete", DiagnosticSeverity::Info)
                } else {
                    counts.pending += 1;
                    ("pending", DiagnosticSeverity::Warning)
                };
                let message = format!(
                    "{} part: {} (content-type={}, parser={})",
                    status,
                    path,
                    ct.unwrap_or("unknown"),
                    spec.expected_parser
                );
                push_entry(
                    entries,
                    severity,
                    "COVERAGE_PART",
                    message.clone(),
                    Some(path),
                );
                push_entry(
                    entries,
                    DiagnosticSeverity::Info,
                    "COVERAGE_PART_ROW",
                    message,
                    Some(path),
                );
            }
        }
        if !matched {
            counts.missing += 1;
            push_entry(
                entries,
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

    counts
}

fn collect_content_type_data(
    content_types: &ContentTypes,
    all_paths: &[String],
) -> ContentTypeInventory {
    let mut inventory = Vec::new();
    let mut unknown_ct_paths = Vec::new();
    let mut ct_histogram = HashMap::new();

    for path in all_paths {
        if let Some(ct) = content_types.get_content_type(path) {
            inventory.push((path.clone(), ct.to_string()));
            *ct_histogram.entry(ct.to_string()).or_insert(0) += 1;
        } else {
            unknown_ct_paths.push(path.clone());
        }
    }
    inventory.sort_by(|a, b| a.0.cmp(&b.0));
    unknown_ct_paths.sort();
    (inventory, unknown_ct_paths, ct_histogram)
}

fn emit_content_type_diagnostics(
    entries: &mut Vec<DiagnosticEntry>,
    inventory: Vec<(String, String)>,
    unknown_ct_paths: &[String],
) {
    for (path, ct) in inventory {
        push_entry(
            entries,
            DiagnosticSeverity::Info,
            "CONTENT_TYPE_INVENTORY",
            format!("content-type: {} => {}", path, ct),
            Some(&path),
        );
    }
    for path in unknown_ct_paths {
        push_entry(
            entries,
            DiagnosticSeverity::Warning,
            "CONTENT_TYPE_UNKNOWN",
            format!("no content-type entry for path {}", path),
            Some(path),
        );
    }
}

fn emit_summary_diagnostics(
    entries: &mut Vec<DiagnosticEntry>,
    counts: &CoverageCounts,
    unknown_count: usize,
) {
    push_entry(
        entries,
        DiagnosticSeverity::Info,
        "COVERAGE_SUMMARY",
        format!(
            "coverage summary: matched={}, complete={}, pending={}, missing_patterns={}, unknown_content_types={}",
            counts.matched_parts, counts.complete, counts.pending, counts.missing, unknown_count
        ),
        None,
    );
    push_entry(
        entries,
        DiagnosticSeverity::Info,
        "COVERAGE_COUNTS",
        format!(
            "counts: complete={}, pending={}, missing_patterns={}, unknown_content_types={}",
            counts.complete, counts.pending, counts.missing, unknown_count
        ),
        None,
    );
}

fn emit_histogram_diagnostics(
    entries: &mut Vec<DiagnosticEntry>,
    ct_histogram: HashMap<String, usize>,
) {
    let mut hist: Vec<(String, usize)> = ct_histogram.into_iter().collect();
    hist.sort_by(|a, b| a.0.cmp(&b.0));
    for (ct, count) in hist {
        push_entry(
            entries,
            DiagnosticSeverity::Info,
            "CONTENT_TYPE_HISTOGRAM",
            format!("content-type-count: {} => {}", ct, count),
            Some(&ct),
        );
    }
}
