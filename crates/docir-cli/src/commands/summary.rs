//! Summary command implementation.

use anyhow::{Context, Result};
use docir_app::ParserConfig;
use docir_core::ir::{IRNode, IrNode};
use docir_core::visitor::{IrVisitor, NodeCounter, PreOrderWalker, VisitControl, VisitorResult};
use std::path::PathBuf;

use crate::commands::util::build_app;

pub fn run(input: PathBuf, parser_config: &ParserConfig) -> Result<()> {
    // Parse the document
    let app = build_app(parser_config);
    let parsed = app
        .parse_file(&input)
        .with_context(|| format!("Failed to parse {}", input.display()))?;

    // Get document info
    let doc = parsed.document().context("Failed to get document root")?;

    // Print header
    println!("Document Summary");
    println!("================");
    println!();

    // Basic info
    println!("File: {}", input.display());
    println!("Format: {}", doc.format.display_name());
    println!();

    // Metadata
    if let Some(meta_id) = doc.metadata {
        if let Some(IRNode::Metadata(meta)) = parsed.store.get(meta_id) {
            println!("Metadata:");
            if let Some(title) = &meta.title {
                println!("  Title: {}", title);
            }
            if let Some(creator) = &meta.creator {
                println!("  Author: {}", creator);
            }
            if let Some(modified) = &meta.modified {
                println!("  Modified: {}", modified);
            }
            if let Some(app) = &meta.application {
                println!("  Application: {}", app);
            }
            println!();
        }
    }

    // Count nodes
    let mut counter = NodeCounter::new();
    let mut walker = PreOrderWalker::new(&parsed.store, parsed.root_id);
    let _ = walker.walk(&mut counter);

    println!("Structure:");
    let mut counts: Vec<_> = counter.counts.iter().collect();
    counts.sort_by_key(|(_, v)| std::cmp::Reverse(*v));
    for (node_type, count) in counts {
        if *count > 0 {
            println!("  {}: {}", node_type, count);
        }
    }
    println!();

    // Text stats
    let mut text_collector = TextStats::new();
    let mut walker = PreOrderWalker::new(&parsed.store, parsed.root_id);
    let _ = walker.walk(&mut text_collector);

    println!("Text Statistics:");
    println!("  Characters: {}", text_collector.char_count);
    println!("  Words: ~{}", text_collector.word_count);
    println!();

    if let Some(metrics) = parsed.metrics.as_ref() {
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
    let security = &doc.security;
    println!("Security:");
    println!("  Threat Level: {}", security.threat_level);
    println!(
        "  VBA Macros: {}",
        if security.macro_project.is_some() {
            "YES - DETECTED"
        } else {
            "No"
        }
    );
    println!(
        "  OLE Objects: {}",
        if security.ole_objects.is_empty() {
            "No".to_string()
        } else {
            format!("{} found", security.ole_objects.len())
        }
    );
    println!(
        "  External References: {}",
        if security.external_refs.is_empty() {
            "No".to_string()
        } else {
            format!("{} found", security.external_refs.len())
        }
    );
    println!(
        "  DDE Fields: {}",
        if security.dde_fields.is_empty() {
            "No".to_string()
        } else {
            format!("{} found", security.dde_fields.len())
        }
    );

    if !security.threat_indicators.is_empty() {
        println!();
        println!("Threat Indicators:");
        for indicator in &security.threat_indicators {
            println!(
                "  [{:?}] {:?}: {}",
                indicator.severity, indicator.indicator_type, indicator.description
            );
        }
    }

    Ok(())
}

/// Visitor that collects text statistics.
struct TextStats {
    char_count: usize,
    word_count: usize,
}

impl TextStats {
    fn new() -> Self {
        Self {
            char_count: 0,
            word_count: 0,
        }
    }
}

impl IrVisitor for TextStats {
    fn visit_run(&mut self, run: &docir_core::ir::Run) -> VisitorResult<VisitControl> {
        self.char_count += run.text.chars().count();
        self.word_count += run.text.split_whitespace().count();
        Ok(VisitControl::Continue)
    }
}
