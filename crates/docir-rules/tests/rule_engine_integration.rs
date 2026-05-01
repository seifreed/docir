use docir_rules::{RuleEngine, RuleProfile, Severity};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

fn finding_count_by_rule(
    report: &docir_rules::RuleReport,
) -> std::collections::BTreeMap<String, usize> {
    let mut map = std::collections::BTreeMap::new();
    for finding in &report.findings {
        *map.entry(finding.rule_id.clone()).or_insert(0) += 1;
    }
    map
}

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(Path::parent)
        .expect("repo root")
        .to_path_buf()
}

fn parse_fixture(relative: &str) -> docir_parser::parser::ParsedDocument {
    let parser = docir_parser::DocumentParser::new();
    parser
        .parse_file(repo_root().join(relative))
        .expect("fixture parse")
}

#[test]
fn default_rules_run_on_rich_fixture_returns_report() {
    let parsed = parse_fixture("fixtures/ooxml/rich.docx");
    let engine = RuleEngine::with_default_rules();
    let report = engine.run(&parsed.store, parsed.root_id);
    assert!(report.findings.iter().all(|f| !f.rule_id.is_empty()));
}

#[test]
fn profile_can_disable_all_rules_via_enabled_list() {
    let parsed = parse_fixture("fixtures/ooxml/rich.xlsx");
    let profile = RuleProfile {
        enabled_rules: Some(Vec::new()),
        disabled_rules: Vec::new(),
        severity_overrides: BTreeMap::new(),
        thresholds: Default::default(),
    };
    let engine = RuleEngine::with_profile(profile.clone());
    let report = engine.run_with_profile(&parsed.store, parsed.root_id, &profile);
    assert!(report.findings.is_empty());
}

#[test]
fn profile_severity_overrides_are_applied_on_real_fixture() {
    let parsed = parse_fixture("fixtures/ooxml/rich.docx");
    let baseline_engine = RuleEngine::with_default_rules();
    let baseline = baseline_engine.run(&parsed.store, parsed.root_id);

    let Some(first_rule_id) = baseline.findings.first().map(|f| f.rule_id.clone()) else {
        // Fixture may legitimately yield no findings; skip strict assertion in that case.
        return;
    };

    let mut overrides = BTreeMap::new();
    overrides.insert(first_rule_id.clone(), Severity::Critical);
    let profile = RuleProfile {
        enabled_rules: None,
        disabled_rules: Vec::new(),
        severity_overrides: overrides,
        thresholds: Default::default(),
    };

    let engine = RuleEngine::with_default_rules();
    let report = engine.run_with_profile(&parsed.store, parsed.root_id, &profile);
    assert!(report
        .findings
        .iter()
        .any(|f| f.rule_id == first_rule_id && f.severity == Severity::Critical));
}

#[test]
fn threshold_burst_rules_trigger_only_when_counts_exceed_thresholds() {
    let parsed = parse_fixture("fixtures/ooxml/rich.xlsx");
    let doc = parsed.document().expect("document root");

    let ole_count = doc.security.ole_objects.len();
    let activex_count = doc.security.activex_controls.len();
    let external_link_count = parsed
        .store
        .values()
        .filter(|node| {
            matches!(
                node,
                docir_core::ir::IRNode::Hyperlink(link) if link.is_external
            )
        })
        .count();

    let strict_thresholds = docir_rules::RuleThresholds {
        max_external_links: Some(external_link_count),
        max_ole_objects: Some(ole_count),
        max_activex_controls: Some(activex_count),
    };

    let strict_profile = RuleProfile {
        enabled_rules: None,
        disabled_rules: Vec::new(),
        severity_overrides: BTreeMap::new(),
        thresholds: strict_thresholds,
    };

    let engine = RuleEngine::with_default_rules();
    let strict_report = engine.run_with_profile(&parsed.store, parsed.root_id, &strict_profile);
    let strict_ids: std::collections::BTreeSet<_> = strict_report
        .findings
        .iter()
        .map(|f| f.rule_id.as_str())
        .collect();
    assert!(!strict_ids.contains("SEC-006"));
    assert!(!strict_ids.contains("SEC-008"));
    assert!(!strict_ids.contains("SEC-012"));

    let low_thresholds = docir_rules::RuleThresholds {
        max_external_links: Some(external_link_count.saturating_sub(1)),
        max_ole_objects: Some(ole_count.saturating_sub(1)),
        max_activex_controls: Some(activex_count.saturating_sub(1)),
    };

    let low_profile = RuleProfile {
        enabled_rules: None,
        disabled_rules: Vec::new(),
        severity_overrides: BTreeMap::new(),
        thresholds: low_thresholds,
    };

    let low_report = engine.run_with_profile(&parsed.store, parsed.root_id, &low_profile);
    let low_ids: std::collections::BTreeSet<_> = low_report
        .findings
        .iter()
        .map(|f| f.rule_id.as_str())
        .collect();

    if ole_count > 0 {
        assert!(low_ids.contains("SEC-006"));
    }
    if activex_count > 0 {
        assert!(low_ids.contains("SEC-008"));
    }
    if external_link_count > 0 {
        assert!(low_ids.contains("SEC-012"));
    }
}

#[test]
fn enabled_and_disabled_filters_apply_with_disabled_precedence() {
    let parsed = parse_fixture("fixtures/ooxml/rich.docx");
    let engine = RuleEngine::with_default_rules();
    let baseline = engine.run(&parsed.store, parsed.root_id);
    let counts = finding_count_by_rule(&baseline);

    let mut present_rules: Vec<String> = counts
        .iter()
        .filter_map(|(id, count)| if *count > 0 { Some(id.clone()) } else { None })
        .collect();
    present_rules.sort();

    let Some(keep_rule) = present_rules.first().cloned() else {
        return;
    };

    let disable_rule = present_rules
        .get(1)
        .cloned()
        .unwrap_or_else(|| keep_rule.clone());

    let profile = RuleProfile {
        enabled_rules: Some(vec![keep_rule.clone(), disable_rule.clone()]),
        disabled_rules: vec![disable_rule.clone()],
        severity_overrides: BTreeMap::new(),
        thresholds: Default::default(),
    };

    let filtered = engine.run_with_profile(&parsed.store, parsed.root_id, &profile);
    let filtered_ids: std::collections::BTreeSet<_> = filtered
        .findings
        .iter()
        .map(|f| f.rule_id.clone())
        .collect();

    assert!(!filtered_ids.contains(&disable_rule));
    if keep_rule != disable_rule {
        assert!(
            filtered_ids.is_empty() || filtered_ids.iter().all(|id| id == &keep_rule),
            "only explicitly enabled non-disabled rule should remain"
        );
    }
}
