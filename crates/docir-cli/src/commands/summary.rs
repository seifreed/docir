//! Summary command implementation.

use anyhow::{Context, Result};
use docir_app::{summarize_document, ParserConfig};
use std::path::PathBuf;

use crate::commands::util::parse_document;

pub fn run(input: PathBuf, parser_config: &ParserConfig) -> Result<()> {
    // Parse the document
    let parsed = parse_document(&input, parser_config)?;

    let summary = summarize_document(&parsed).context("Failed to build document summary")?;

    // Print header
    println!("Document Summary");
    println!("================");
    println!();

    // Basic info
    println!("File: {}", input.display());
    println!("Format: {}", summary.format);
    println!();

    // Metadata
    if summary.metadata.title.is_some()
        || summary.metadata.author.is_some()
        || summary.metadata.modified.is_some()
        || summary.metadata.application.is_some()
    {
        println!("Metadata:");
        if let Some(title) = &summary.metadata.title {
            println!("  Title: {}", title);
        }
        if let Some(creator) = &summary.metadata.author {
            println!("  Author: {}", creator);
        }
        if let Some(modified) = &summary.metadata.modified {
            println!("  Modified: {}", modified);
        }
        if let Some(app) = &summary.metadata.application {
            println!("  Application: {}", app);
        }
        println!();
    }

    println!("Structure:");
    for count in &summary.node_counts {
        if count.count > 0 {
            println!("  {}: {}", count.node_type, count.count);
        }
    }
    println!();

    println!("Text Statistics:");
    println!("  Characters: {}", summary.text_stats.char_count);
    println!("  Words: ~{}", summary.text_stats.word_count);
    println!();

    if let Some(metrics) = &summary.metrics {
        println!("Parse Metrics (ms):");
        println!("  Content Types: {}", metrics.content_types_ms);
        println!("  Relationships: {}", metrics.relationships_ms);
        println!("  Main Parse: {}", metrics.main_parse_ms);
        println!("  Shared Parts: {}", metrics.shared_parts_ms);
        println!("  Security Scan: {}", metrics.security_scan_ms);
        println!("  Extension Parts: {}", metrics.extension_parts_ms);
        println!("  Normalization: {}", metrics.normalization_ms);
        println!();
    }

    // Security summary
    println!("Security:");
    println!("  Threat Level: {}", summary.security.threat_level);
    println!(
        "  VBA Macros: {}",
        if summary.security.has_macro_project {
            "YES - DETECTED"
        } else {
            "No"
        }
    );
    println!(
        "  OLE Objects: {}",
        if summary.security.ole_objects == 0 {
            "No".to_string()
        } else {
            format!("{} found", summary.security.ole_objects)
        }
    );
    println!(
        "  External References: {}",
        if summary.security.external_refs == 0 {
            "No".to_string()
        } else {
            format!("{} found", summary.security.external_refs)
        }
    );
    println!(
        "  DDE Fields: {}",
        if summary.security.dde_fields == 0 {
            "No".to_string()
        } else {
            format!("{} found", summary.security.dde_fields)
        }
    );

    if !summary.threat_indicators.is_empty() {
        println!();
        println!("Threat Indicators:");
        for indicator in &summary.threat_indicators {
            println!(
                "  [{}] {}: {}",
                indicator.severity, indicator.indicator_type, indicator.description
            );
        }
    }

    Ok(())
}
