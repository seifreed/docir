//! Coverage command implementation.

use crate::commands::util::parse_document;
use crate::{CoverageExportFormat, CoverageExportMode};
use anyhow::{Context, Result};
use docir_app::ParserConfig;
use docir_core::ir::{DiagnosticEntry, DiagnosticSeverity, IRNode};
use serde::Serialize;
use std::path::PathBuf;

#[derive(Debug, Clone)]
pub struct CoverageOptions {
    pub json: bool,
    pub details: bool,
    pub inventory: bool,
    pub unknown: bool,
    pub export: Option<PathBuf>,
    pub export_format: CoverageExportFormat,
    pub export_mode: CoverageExportMode,
}

#[derive(Debug, Serialize)]
struct CoverageReport {
    summary: Option<String>,
    unparsed_summary: Option<String>,
    counts: Option<String>,
    parts: Vec<DiagnosticEntry>,
    part_rows: Vec<DiagnosticEntry>,
    missing: Vec<DiagnosticEntry>,
    inventory: Vec<DiagnosticEntry>,
    unknown: Vec<DiagnosticEntry>,
    histogram: Vec<DiagnosticEntry>,
}

pub fn run(input: PathBuf, options: CoverageOptions, parser_config: &ParserConfig) -> Result<()> {
    let parsed = parse_document(&input, parser_config)?;
    let doc = parsed.document().context("Failed to get document root")?;
    let entries = collect_coverage_entries(&parsed, doc);
    let report = build_report(entries);

    if options.json {
        output_json_report(&report, options.export.as_ref(), &options)?;
        return Ok(());
    }

    if let Some(path) = options.export.as_ref() {
        let json = serde_json::to_string_pretty(&report)?;
        write_export(
            path,
            options.export_format,
            options.export_mode,
            &report,
            json,
        )?;
    }

    print_text_report(&input, doc.format.display_name(), &report, &options);

    Ok(())
}

fn collect_coverage_entries(
    parsed: &docir_app::ParsedDocument,
    doc: &docir_core::ir::Document,
) -> Vec<DiagnosticEntry> {
    let mut entries = Vec::new();
    for diag_id in &doc.diagnostics {
        if let Some(IRNode::Diagnostics(diag)) = parsed.store().get(*diag_id) {
            entries.extend(diag.entries.clone());
        }
    }
    entries
}

fn build_report(entries: Vec<DiagnosticEntry>) -> CoverageReport {
    let mut report = CoverageReport {
        summary: None,
        unparsed_summary: None,
        counts: None,
        parts: Vec::new(),
        part_rows: Vec::new(),
        missing: Vec::new(),
        inventory: Vec::new(),
        unknown: Vec::new(),
        histogram: Vec::new(),
    };

    for entry in entries {
        match entry.code.as_str() {
            "COVERAGE_SUMMARY" => report.summary = Some(entry.message),
            "UNPARSED_SUMMARY" => report.unparsed_summary = Some(entry.message),
            "COVERAGE_COUNTS" => report.counts = Some(entry.message),
            "COVERAGE_PART" => report.parts.push(entry),
            "COVERAGE_PART_ROW" => report.part_rows.push(entry),
            "COVERAGE_MISSING" => report.missing.push(entry),
            "CONTENT_TYPE_INVENTORY" => report.inventory.push(entry),
            "CONTENT_TYPE_UNKNOWN" => report.unknown.push(entry),
            "CONTENT_TYPE_HISTOGRAM" => report.histogram.push(entry),
            _ => {}
        }
    }

    report
}

fn output_json_report(
    report: &CoverageReport,
    export_path: Option<&PathBuf>,
    options: &CoverageOptions,
) -> Result<()> {
    let json = serde_json::to_string_pretty(report)?;
    println!("{}", json);
    if let Some(path) = export_path {
        write_export(
            path,
            options.export_format,
            options.export_mode,
            report,
            json,
        )?;
    }
    Ok(())
}

fn print_text_report(
    input: &PathBuf,
    format_name: &str,
    report: &CoverageReport,
    options: &CoverageOptions,
) {
    println!("Coverage Report");
    println!("===============");
    println!();
    println!("File: {}", input.display());
    println!("Format: {}", format_name);
    println!();

    match &report.summary {
        Some(summary) => println!("{}", summary),
        None => println!("coverage summary: missing"),
    }
    if let Some(unparsed) = &report.unparsed_summary {
        println!("{}", unparsed);
    }
    if let Some(counts) = &report.counts {
        println!("{}", counts);
    }

    let complete = report
        .parts
        .iter()
        .filter(|e| matches!(e.severity, DiagnosticSeverity::Info))
        .count();
    let pending = report
        .parts
        .iter()
        .filter(|e| matches!(e.severity, DiagnosticSeverity::Warning))
        .count();
    println!(
        "parts: complete={}, pending={}, missing_patterns={}, unknown_content_types={}",
        complete,
        pending,
        report.missing.len(),
        report.unknown.len()
    );

    if options.details {
        print_entry_section("Matched Parts", &report.parts);
        print_entry_section("Missing Patterns", &report.missing);
    }
    if options.inventory {
        print_entry_section("Content Types Inventory", &report.inventory);
    }
    print_entry_section("Content Type Histogram", &report.histogram);
    if options.unknown {
        print_entry_section("Unknown Content Types", &report.unknown);
    }
}

fn print_entry_section(label: &str, entries: &[DiagnosticEntry]) {
    if entries.is_empty() {
        return;
    }
    println!();
    println!("{label}:");
    let mut rows = entries.to_vec();
    rows.sort_by(|a, b| a.path.cmp(&b.path));
    for entry in rows {
        println!("  {}", entry.message);
    }
}

fn write_export(
    path: &PathBuf,
    format: CoverageExportFormat,
    mode: CoverageExportMode,
    report: &CoverageReport,
    json: String,
) -> Result<()> {
    match format {
        CoverageExportFormat::Json => {
            let payload = match mode {
                CoverageExportMode::Full => json,
                CoverageExportMode::Parts => {
                    let rows = build_part_rows(report);
                    serde_json::to_string_pretty(&rows)?
                }
            };
            std::fs::write(path, payload)?;
        }
        CoverageExportFormat::Csv => {
            let csv = match mode {
                CoverageExportMode::Full => render_csv(report),
                CoverageExportMode::Parts => render_parts_csv(report),
            };
            std::fs::write(path, csv)?;
        }
    }
    Ok(())
}

fn render_csv(report: &CoverageReport) -> String {
    let mut out = String::new();
    out.push_str("section,code,severity,message,path\n");
    write_csv_entries(&mut out, "parts", &report.parts);
    write_csv_entries(&mut out, "part_rows", &report.part_rows);
    write_csv_entries(&mut out, "missing", &report.missing);
    write_csv_entries(&mut out, "inventory", &report.inventory);
    write_csv_entries(&mut out, "unknown", &report.unknown);
    write_csv_entries(&mut out, "histogram", &report.histogram);
    if let Some(summary) = &report.summary {
        let msg = escape_csv(summary);
        out.push_str(&format!("summary,COVERAGE_SUMMARY,Info,{msg},\n"));
    }
    out
}

fn render_parts_csv(report: &CoverageReport) -> String {
    let mut out = String::new();
    out.push_str("status,path,content_type,parser\n");
    for entry in &report.part_rows {
        if let Some(row) = parse_part_row(&entry.message) {
            out.push_str(&format!(
                "{},{},{},{}\n",
                escape_csv(&row.status),
                escape_csv(&row.path),
                escape_csv(&row.content_type),
                escape_csv(&row.parser)
            ));
        }
    }
    out
}

#[derive(Debug, Serialize)]
struct PartRow {
    status: String,
    path: String,
    content_type: String,
    parser: String,
}

fn build_part_rows(report: &CoverageReport) -> Vec<PartRow> {
    let mut rows = Vec::new();
    for entry in &report.part_rows {
        if let Some(row) = parse_part_row(&entry.message) {
            rows.push(row);
        }
    }
    rows
}

fn parse_part_row(message: &str) -> Option<PartRow> {
    let (status_part, rest) = message.split_once(" part: ")?;
    let (path_part, tail) = rest.split_once(" (content-type=")?;
    let (ct_part, parser_part) = tail.split_once(", parser=")?;
    let parser = parser_part.trim_end_matches(')').to_string();
    Some(PartRow {
        status: status_part.to_string(),
        path: path_part.to_string(),
        content_type: ct_part.to_string(),
        parser,
    })
}

fn write_csv_entries(out: &mut String, section: &str, entries: &[DiagnosticEntry]) {
    for entry in entries {
        let code = escape_csv(&entry.code);
        let severity = match entry.severity {
            DiagnosticSeverity::Info => "Info",
            DiagnosticSeverity::Warning => "Warning",
            DiagnosticSeverity::Error => "Error",
        };
        let message = escape_csv(&entry.message);
        let path = entry
            .path
            .as_ref()
            .map(|p| escape_csv(p))
            .unwrap_or_else(|| "".to_string());
        out.push_str(&format!("{section},{code},{severity},{message},{path}\n"));
    }
}

fn escape_csv(value: &str) -> String {
    if value.contains(',') || value.contains('"') || value.contains('\n') {
        let escaped = value.replace('"', "\"\"");
        format!("\"{escaped}\"")
    } else {
        value.to_string()
    }
}
