//! Summary command implementation.

use anyhow::{Context, Result};
use docir_app::{summarize_document, ParserConfig};
use std::path::PathBuf;

use crate::commands::util::parse_document;

pub fn run(input: PathBuf, parser_config: &ParserConfig) -> Result<()> {
    let parsed = parse_document(&input, parser_config)?;
    let summary = summarize_document(&parsed).context("Failed to build document summary")?;

    println!("Document Summary");
    println!("================");
    println!();
    println!("File: {}", input.display());
    println!("Format: {}", summary.format);
    println!();

    print_metadata(&summary);
    print_structure(&summary);
    print_text_stats(&summary);
    print_metrics(&summary);
    print_security(&summary);
    print_threat_indicators(&summary);

    Ok(())
}

fn print_metadata(summary: &docir_app::DocumentSummary) {
    if summary.metadata.title.is_none()
        && summary.metadata.author.is_none()
        && summary.metadata.modified.is_none()
        && summary.metadata.application.is_none()
    {
        return;
    }

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

fn print_structure(summary: &docir_app::DocumentSummary) {
    println!("Structure:");
    for count in &summary.node_counts {
        if count.count > 0 {
            println!("  {}: {}", count.node_type, count.count);
        }
    }
    println!();
}

fn print_text_stats(summary: &docir_app::DocumentSummary) {
    println!("Text Statistics:");
    println!("  Characters: {}", summary.text_stats.char_count);
    println!("  Words: ~{}", summary.text_stats.word_count);
    println!();
}

fn print_metrics(summary: &docir_app::DocumentSummary) {
    let Some(metrics) = &summary.metrics else {
        return;
    };

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

fn print_security(summary: &docir_app::DocumentSummary) {
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
        count_or_none(summary.security.ole_objects)
    );
    println!(
        "  External References: {}",
        count_or_none(summary.security.external_refs)
    );
    println!(
        "  DDE Fields: {}",
        count_or_none(summary.security.dde_fields)
    );
}

fn count_or_none(count: usize) -> String {
    if count == 0 {
        "No".to_string()
    } else {
        format!("{} found", count)
    }
}

fn print_threat_indicators(summary: &docir_app::DocumentSummary) {
    if summary.threat_indicators.is_empty() {
        return;
    }

    println!();
    println!("Threat Indicators:");
    for indicator in &summary.threat_indicators {
        println!(
            "  [{}] {}: {}",
            indicator.severity, indicator.indicator_type, indicator.description
        );
    }
}
