//! Analyst-facing document indicator scorecard.

use anyhow::{Context, Result};
use docir_app::{IndicatorReport, ParserConfig};
use std::fs;
use std::path::PathBuf;

use crate::commands::util::{build_app, push_bullet_line, push_labeled_line, run_dual_output};

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let app = build_app(parser_config);
    let source_bytes =
        fs::read(&input).with_context(|| format!("Failed to read {}", input.display()))?;
    let parsed = app
        .parse_bytes(&source_bytes)
        .with_context(|| format!("Failed to parse {}", input.display()))?;
    let report = app.build_indicator_report_with_bytes(&parsed, &source_bytes);
    run_dual_output(&report, "report", json, pretty, output, format_report_text)
}

fn format_report_text(report: &IndicatorReport) -> String {
    let directory_score = report
        .indicators
        .iter()
        .find(|indicator| indicator.key == "cfb-directory-score")
        .map(|indicator| indicator.value.clone())
        .unwrap_or_else(|| "n/a".to_string());
    let sector_score = report
        .indicators
        .iter()
        .find(|indicator| indicator.key == "cfb-sector-score")
        .map(|indicator| indicator.value.clone())
        .unwrap_or_else(|| "n/a".to_string());
    let stream_score = report
        .indicators
        .iter()
        .find(|indicator| indicator.key == "cfb-stream-score")
        .map(|indicator| indicator.value.clone())
        .unwrap_or_else(|| "n/a".to_string());
    let dominant_class = report
        .indicators
        .iter()
        .find(|indicator| indicator.key == "cfb-dominant-anomaly-class")
        .map(|indicator| indicator.value.clone())
        .unwrap_or_else(|| "n/a".to_string());

    let mut out = String::new();
    push_labeled_line(&mut out, 0, "Format", &report.document_format);
    push_labeled_line(&mut out, 0, "Container", &report.container);
    push_labeled_line(&mut out, 0, "Overall Risk", report.overall_risk);
    push_labeled_line(&mut out, 0, "Directory Score", &directory_score);
    push_labeled_line(&mut out, 0, "Sector Score", &sector_score);
    push_labeled_line(&mut out, 0, "Stream Score", &stream_score);
    push_labeled_line(&mut out, 0, "Dominant Anomaly", &dominant_class);
    push_labeled_line(&mut out, 0, "Indicators", report.indicator_count);
    out.push_str("\nIndicator Scorecard:\n");
    for indicator in &report.indicators {
        push_bullet_line(
            &mut out,
            2,
            &indicator.key,
            format!("{} [{}]", indicator.value, indicator.risk),
        );
        push_labeled_line(&mut out, 4, "Reason", &indicator.reason);
        if !indicator.evidence.is_empty() {
            push_labeled_line(&mut out, 4, "Evidence", indicator.evidence.join(" | "));
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::{format_report_text, run};
    use crate::test_support;
    use docir_app::{
        inspect_directory_bytes,
        test_support::{build_test_cfb, patch_test_cfb_directory_entry, TestCfbDirectoryPatch},
        DocumentIndicator, IndicatorReport, ParserConfig,
    };
    use docir_core::security::ThreatLevel;
    use serde_json::Value;
    use std::fs;

    #[test]
    fn report_indicators_run_writes_json() {
        let input = test_support::temp_file("legacy", "doc");
        let output = test_support::temp_file("legacy", "json");
        let base = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="LegacyProject"
Module=Module1
"#,
            ),
            ("VBA/Module1", b"Sub AutoOpen()\nEnd Sub\n"),
            ("ObjectPool/1/Ole10Native", b"payload"),
        ]);
        let inspection = inspect_directory_bytes(&base).expect("directory");
        let word = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "WordDocument")
            .expect("word");
        let vba = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "VBA/PROJECT")
            .expect("vba");
        let objectpool = inspection
            .entries
            .iter()
            .find(|entry| entry.path == "ObjectPool/1/Ole10Native")
            .expect("objectpool");
        let patched = patch_test_cfb_directory_entry(
            &patch_test_cfb_directory_entry(
                &patch_test_cfb_directory_entry(
                    &base,
                    vba.entry_index,
                    TestCfbDirectoryPatch {
                        start_sector: Some(word.start_sector),
                        ..Default::default()
                    },
                ),
                objectpool.entry_index,
                TestCfbDirectoryPatch {
                    start_sector: Some(99),
                    ..Default::default()
                },
            ),
            word.entry_index,
            TestCfbDirectoryPatch {
                start_sector: Some(98),
                ..Default::default()
            },
        );
        fs::write(&input, patched).expect("fixture");

        run(
            input.clone(),
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("report-indicators json");

        let text = fs::read_to_string(&output).expect("output");
        let json: Value = serde_json::from_str(&text).expect("valid json");
        let indicators = json["report"]["indicators"]
            .as_array()
            .expect("indicator array");
        let find_indicator = |key: &str| {
            indicators
                .iter()
                .find(|entry| entry["key"].as_str() == Some(key))
                .expect("indicator present")
        };
        assert!(text.contains("\"report\""));
        assert!(text.contains("\"key\": \"macros\""));
        assert!(text.contains("\"key\": \"object-pool\""));
        assert!(text.contains("\"key\": \"cfb-directory-score\""));
        assert!(text.contains("\"key\": \"cfb-sector-score\""));
        assert!(text.contains("\"key\": \"cfb-stream-score\""));
        assert!(text.contains("\"key\": \"cfb-objectpool-corruption\""));
        assert!(text.contains("\"key\": \"cfb-vba-structure-anomalies\""));
        assert!(text.contains("\"key\": \"cfb-main-stream-corruption\""));
        assert!(text.contains("\"key\": \"cfb-dominant-anomaly-class\""));
        assert!(text.contains("objectpool:"));
        assert!(text.contains("vba:"));
        assert!(text.contains("main-stream:word:"));
        assert_eq!(
            find_indicator("cfb-dominant-anomaly-class")["value"]
                .as_str()
                .expect("dominant value"),
            "invalid-start"
        );
        assert_eq!(
            find_indicator("cfb-stream-score")["value"]
                .as_str()
                .expect("stream score"),
            "high"
        );

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn format_report_text_includes_reason_and_evidence() {
        let report = IndicatorReport {
            document_format: "doc".to_string(),
            container: "cfb-ole".to_string(),
            overall_risk: ThreatLevel::Critical,
            indicator_count: 5,
            indicators: vec![
                DocumentIndicator {
                    key: "macros".to_string(),
                    value: "1".to_string(),
                    risk: ThreatLevel::Critical,
                    reason: "VBA project or macro-capable content detected".to_string(),
                    evidence: vec!["VBA".to_string()],
                },
                DocumentIndicator {
                    key: "object-pool".to_string(),
                    value: "1".to_string(),
                    risk: ThreatLevel::High,
                    reason: "Legacy ObjectPool storage entries detected".to_string(),
                    evidence: vec!["ObjectPool/1/Ole10Native".to_string()],
                },
                DocumentIndicator {
                    key: "cfb-directory-score".to_string(),
                    value: "medium".to_string(),
                    risk: ThreatLevel::Medium,
                    reason: "Aggregated directory-graph corruption score for the CFB container"
                        .to_string(),
                    evidence: vec!["directory:cycle:mixed-2-cycle=1".to_string()],
                },
                DocumentIndicator {
                    key: "cfb-sector-score".to_string(),
                    value: "high".to_string(),
                    risk: ThreatLevel::High,
                    reason: "Aggregated sector-allocation corruption score for the CFB container"
                        .to_string(),
                    evidence: vec!["sector:shared-sector:0=WordDocument,VBA/PROJECT".to_string()],
                },
                DocumentIndicator {
                    key: "cfb-stream-score".to_string(),
                    value: "medium".to_string(),
                    risk: ThreatLevel::Medium,
                    reason: "Aggregated stream-level corruption score for the CFB container"
                        .to_string(),
                    evidence: vec!["health:shared:root:WordDocument=1".to_string()],
                },
                DocumentIndicator {
                    key: "cfb-dominant-anomaly-class".to_string(),
                    value: "shared-sector".to_string(),
                    risk: ThreatLevel::Medium,
                    reason: "Dominant low-level anomaly class across the CFB container".to_string(),
                    evidence: Vec::new(),
                },
            ],
        };

        let text = format_report_text(&report);
        assert!(text.contains("Overall Risk: CRITICAL"));
        assert!(text.contains("Directory Score: medium"));
        assert!(text.contains("Sector Score: high"));
        assert!(text.contains("Stream Score: medium"));
        assert!(text.contains("Dominant Anomaly: shared-sector"));
        assert!(text.contains("- macros: 1 [CRITICAL]"));
        assert!(text.contains("Reason: VBA project or macro-capable content detected"));
        assert!(text.contains("Evidence: ObjectPool/1/Ole10Native"));
        assert!(text.contains("- cfb-directory-score: medium [MEDIUM]"));
    }
}
