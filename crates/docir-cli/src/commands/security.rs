//! Security command implementation.

use crate::commands::util::build_parser;
use anyhow::{Context, Result};
use docir_parser::ParserConfig;
use docir_security::SecurityAnalyzer;
use std::path::PathBuf;

pub fn run(input: PathBuf, json: bool, verbose: bool, parser_config: &ParserConfig) -> Result<()> {
    // Parse the document
    let parser = build_parser(parser_config);
    let parsed = parser
        .parse_file(&input)
        .with_context(|| format!("Failed to parse {}", input.display()))?;

    // Run security analysis
    let mut analyzer = SecurityAnalyzer::new();
    let result = analyzer.analyze(&parsed.store, parsed.root_id);

    if json {
        // Output as JSON
        let output = serde_json::json!({
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
        println!("{}", serde_json::to_string_pretty(&output)?);
    } else {
        // Human-readable output
        println!("Security Analysis Report");
        println!("========================");
        println!();
        println!("File: {}", input.display());
        println!();

        // Threat level with color indication (using ASCII)
        let level_indicator = match result.threat_level {
            docir_core::security::ThreatLevel::None => "[OK]    ",
            docir_core::security::ThreatLevel::Low => "[LOW]   ",
            docir_core::security::ThreatLevel::Medium => "[MEDIUM]",
            docir_core::security::ThreatLevel::High => "[HIGH]  ",
            docir_core::security::ThreatLevel::Critical => "[CRIT]  ",
        };
        println!("Threat Level: {} {}", level_indicator, result.threat_level);
        println!();

        // Feature detection
        println!("Security Features Detected:");
        print_feature("VBA Macros", result.has_macros);
        print_feature("OLE Objects", result.has_ole_objects);
        print_feature("External References", result.has_external_refs);
        print_feature("DDE Fields", result.has_dde);
        print_feature("XLM Macros (Excel 4.0)", result.has_xlm_macros);
        println!();

        // Findings
        if result.findings.is_empty() {
            println!("No specific security findings.");
        } else {
            println!("Findings ({}):", result.findings.len());
            println!();

            for (i, finding) in result.findings.iter().enumerate() {
                let severity_marker = match finding.severity {
                    docir_core::security::ThreatLevel::None => "[ ]",
                    docir_core::security::ThreatLevel::Low => "[!]",
                    docir_core::security::ThreatLevel::Medium => "[!!]",
                    docir_core::security::ThreatLevel::High => "[!!!]",
                    docir_core::security::ThreatLevel::Critical => "[!!!!]",
                };

                println!(
                    "  {}. {} {:?}",
                    i + 1,
                    severity_marker,
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

        // Recommendation
        if result.has_concerns() {
            println!("---");
            println!("RECOMMENDATION: This document contains potentially dangerous content.");
            println!("Exercise caution before opening or enabling macros.");
        }
    }

    Ok(())
}

fn print_feature(name: &str, detected: bool) {
    let marker = if detected { "[X]" } else { "[ ]" };
    let status = if detected { "DETECTED" } else { "Not detected" };
    println!("  {} {}: {}", marker, name, status);
}
