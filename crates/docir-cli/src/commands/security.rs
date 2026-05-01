//! Security command implementation.

use anyhow::Result;
use docir_app::{AnalysisResult, ParserConfig};
use serde_json::json;
use std::path::{Path, PathBuf};

use crate::commands::util::{build_app_and_parse, write_json_output};

/// Public API entrypoint: run.
pub fn run(input: PathBuf, json: bool, verbose: bool, parser_config: &ParserConfig) -> Result<()> {
    let (app, parsed) = build_app_and_parse(&input, parser_config)?;
    let result = app.analyze_security(&parsed);

    if json {
        let output = build_json_report(&input, &result);
        write_json_output(&output, true, None)?;
    } else {
        print_human_report(&input, &result, verbose);
    }

    Ok(())
}

fn build_json_report(input: &Path, result: &AnalysisResult) -> serde_json::Value {
    let output = json!({
        "file": input.display().to_string(),
        "threat_level": format!("{}", result.threat_level),
        "has_macros": result.has_macros,
        "has_ole_objects": result.has_ole_objects,
        "has_external_refs": result.has_external_refs,
        "has_dde": result.has_dde,
        "has_xlm_macros": result.has_xlm_macros,
        "findings_count": result.findings.len(),
        "findings": result.findings.iter().map(|f| {
            serde_json::json!({
                "type": format!("{:?}", f.indicator_type),
                "severity": format!("{}", f.severity),
                "description": f.description,
                "location": f.location,
            })
        }).collect::<Vec<_>>(),
    });
    output
}

fn print_human_report(input: &Path, result: &AnalysisResult, verbose: bool) {
    println!("Security Analysis Report");
    println!("========================");
    println!();
    println!("File: {}", input.display());
    println!();
    println!(
        "Threat Level: {} {}",
        threat_level_marker(result.threat_level),
        result.threat_level
    );
    println!();

    println!("Security Features Detected:");
    print_feature("VBA Macros", result.has_macros);
    print_feature("OLE Objects", result.has_ole_objects);
    print_feature("External References", result.has_external_refs);
    print_feature("DDE Fields", result.has_dde);
    print_feature("XLM Macros (Excel 4.0)", result.has_xlm_macros);
    println!();

    if result.findings.is_empty() {
        println!("No specific security findings.");
    } else {
        print_findings(result, verbose);
    }

    if result.has_concerns() {
        println!("---");
        println!("RECOMMENDATION: This document contains potentially dangerous content.");
        println!("Exercise caution before opening or enabling macros.");
    }
}

fn print_findings(result: &AnalysisResult, verbose: bool) {
    println!("Findings ({}):", result.findings.len());
    println!();

    for (i, finding) in result.findings.iter().enumerate() {
        println!(
            "  {}. {} {:?}",
            i + 1,
            severity_marker(finding.severity),
            finding.indicator_type
        );
        println!("     {}", finding.description);

        if verbose {
            if let Some(loc) = &finding.location {
                println!("     Location: {}", loc);
            }
            if let Some(node_id) = &finding.node_id {
                println!("     Node: {}", node_id);
            }
        }
        println!();
    }
}

fn print_feature(name: &str, detected: bool) {
    let marker = if detected { "[X]" } else { "[ ]" };
    let status = if detected { "DETECTED" } else { "Not detected" };
    println!("  {} {}: {}", marker, name, status);
}

fn threat_level_marker(level: docir_core::security::ThreatLevel) -> &'static str {
    match level {
        docir_core::security::ThreatLevel::None => "[OK]    ",
        docir_core::security::ThreatLevel::Low => "[LOW]   ",
        docir_core::security::ThreatLevel::Medium => "[MEDIUM]",
        docir_core::security::ThreatLevel::High => "[HIGH]  ",
        docir_core::security::ThreatLevel::Critical => "[CRIT]  ",
    }
}

fn severity_marker(level: docir_core::security::ThreatLevel) -> &'static str {
    match level {
        docir_core::security::ThreatLevel::None => "[ ]",
        docir_core::security::ThreatLevel::Low => "[!]",
        docir_core::security::ThreatLevel::Medium => "[!!]",
        docir_core::security::ThreatLevel::High => "[!!!]",
        docir_core::security::ThreatLevel::Critical => "[!!!!]",
    }
}

#[cfg(test)]
mod tests {
    use super::{severity_marker, threat_level_marker};
    use docir_core::security::ThreatLevel;

    #[test]
    fn threat_level_markers_cover_all_levels() {
        assert_eq!(threat_level_marker(ThreatLevel::None), "[OK]    ");
        assert_eq!(threat_level_marker(ThreatLevel::Low), "[LOW]   ");
        assert_eq!(threat_level_marker(ThreatLevel::Medium), "[MEDIUM]");
        assert_eq!(threat_level_marker(ThreatLevel::High), "[HIGH]  ");
        assert_eq!(threat_level_marker(ThreatLevel::Critical), "[CRIT]  ");
    }

    #[test]
    fn severity_markers_cover_all_levels() {
        assert_eq!(severity_marker(ThreatLevel::None), "[ ]");
        assert_eq!(severity_marker(ThreatLevel::Low), "[!]");
        assert_eq!(severity_marker(ThreatLevel::Medium), "[!!]");
        assert_eq!(severity_marker(ThreatLevel::High), "[!!!]");
        assert_eq!(severity_marker(ThreatLevel::Critical), "[!!!!]");
    }
}
