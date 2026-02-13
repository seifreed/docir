//! CLI types and configuration wiring.

use clap::{Parser, Subcommand, ValueEnum};
use docir_app::ParserConfig;
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "docir")]
#[command(author = "Marc Rivero López")]
#[command(version)]
#[command(about = "Document Intermediate Representation toolkit for Office documents")]
#[command(long_about = r#"
docir - Document IR for Microsoft Office formats (DOCX, XLSX, PPTX)

A security-focused toolkit for parsing, analyzing, and transforming
Office documents into a semantic Intermediate Representation (IR).

Think of this as "LLVM IR for documents" - providing a structured,
typed, and navigable representation for security analysis, diffing,
and AI consumption.
"#)]
pub(crate) struct Cli {
    /// Maximum total uncompressed size across ZIP entries (bytes)
    #[arg(long, global = true, value_name = "BYTES")]
    pub(crate) zip_max_total_size: Option<u64>,

    /// Maximum size per ZIP entry (bytes)
    #[arg(long, global = true, value_name = "BYTES")]
    pub(crate) zip_max_file_size: Option<u64>,

    /// Maximum number of files in ZIP
    #[arg(long, global = true, value_name = "COUNT")]
    pub(crate) zip_max_file_count: Option<usize>,

    /// Maximum compression ratio for ZIP entries
    #[arg(long, global = true, value_name = "RATIO")]
    pub(crate) zip_max_compression_ratio: Option<f64>,

    /// Maximum path depth inside ZIP
    #[arg(long, global = true, value_name = "DEPTH")]
    pub(crate) zip_max_path_depth: Option<usize>,

    /// Maximum input size for parser entrypoints (bytes)
    #[arg(long, global = true, value_name = "BYTES")]
    pub(crate) max_input_size: Option<u64>,

    /// Force ODF fast mode (skip full cell expansion for large spreadsheets)
    #[arg(long, global = true)]
    pub(crate) odf_fast: bool,

    /// ODF fast-mode threshold for content.xml (bytes)
    #[arg(long, global = true, value_name = "BYTES")]
    pub(crate) odf_fast_threshold_bytes: Option<u64>,

    /// ODF fast-mode sample rows (0 = no sampling)
    #[arg(long, global = true, value_name = "ROWS")]
    pub(crate) odf_fast_sample_rows: Option<u32>,

    /// ODF fast-mode sample columns (0 = no sampling)
    #[arg(long, global = true, value_name = "COLS")]
    pub(crate) odf_fast_sample_cols: Option<u32>,

    /// ODF maximum cells to parse (0 = unlimited)
    #[arg(long, global = true, value_name = "COUNT")]
    pub(crate) odf_max_cells: Option<u64>,

    /// ODF maximum rows to parse (0 = unlimited)
    #[arg(long, global = true, value_name = "COUNT")]
    pub(crate) odf_max_rows: Option<u64>,

    /// ODF maximum paragraphs to parse (0 = unlimited)
    #[arg(long, global = true, value_name = "COUNT")]
    pub(crate) odf_max_paragraphs: Option<u64>,

    /// ODF maximum content.xml bytes (0 = unlimited)
    #[arg(long, global = true, value_name = "BYTES")]
    pub(crate) odf_max_bytes: Option<u64>,

    /// Enable parallel ODF sheet parsing when possible
    #[arg(long, global = true)]
    pub(crate) odf_parallel_sheets: bool,

    /// Max threads for parallel ODF sheet parsing
    #[arg(long, global = true, value_name = "COUNT")]
    pub(crate) odf_parallel_max_threads: Option<usize>,

    /// Password for decrypting encrypted ODF parts
    #[arg(long, global = true, value_name = "PASSWORD")]
    pub(crate) odf_password: Option<String>,

    /// Force parse encrypted HWP streams
    #[arg(long, global = true)]
    pub(crate) hwp_force_parse_encrypted: bool,

    /// Password for decrypting encrypted HWP streams
    #[arg(long, global = true, value_name = "PASSWORD")]
    pub(crate) hwp_password: Option<String>,

    /// Dump HWP stream metadata (hash, size, compression)
    #[arg(long, global = true)]
    pub(crate) hwp_dump_streams: bool,

    /// Enable parser timing metrics
    #[arg(long, global = true)]
    pub(crate) metrics: bool,

    #[command(subcommand)]
    pub(crate) command: Commands,
}

#[derive(Subcommand)]
pub(crate) enum Commands {
    /// Parse a document and output its IR
    Parse {
        /// Path to the OOXML file
        input: PathBuf,

        /// Output format
        #[arg(long, short, default_value = "json")]
        format: OutputFormat,

        /// Pretty-print output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Display a high-level summary of the document
    Summary {
        /// Path to the OOXML file
        input: PathBuf,
    },

    /// Report parser coverage for the document
    Coverage {
        /// Path to the OOXML file
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Include per-part coverage details
        #[arg(long)]
        details: bool,

        /// Include content-type inventory
        #[arg(long)]
        inventory: bool,

        /// Include paths with unknown content-types
        #[arg(long)]
        unknown: bool,

        /// Export coverage report JSON to a file
        #[arg(long)]
        export: Option<PathBuf>,

        /// Export format (json or csv)
        #[arg(long, default_value = "json")]
        export_format: CoverageExportFormat,

        /// Export mode (full report or parts-only)
        #[arg(long, default_value = "full")]
        export_mode: CoverageExportMode,
    },

    /// Perform security analysis on the document
    Security {
        /// Path to the OOXML file
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Verbose output with all findings
        #[arg(long, short)]
        verbose: bool,
    },

    /// Dump a specific node from the IR by ID
    #[command(name = "dump-node")]
    DumpNode {
        /// Path to the OOXML file
        input: PathBuf,

        /// Node ID to dump
        #[arg(long)]
        node_id: String,

        /// Output format
        #[arg(long, short, default_value = "json")]
        format: OutputFormat,
    },

    /// Diff two documents and output the IR diff
    Diff {
        /// Path to the left (base) OOXML file
        left: PathBuf,

        /// Path to the right (compare) OOXML file
        right: PathBuf,

        /// Pretty-print output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Run rule engine on a document
    Rules {
        /// Path to the OOXML file
        input: PathBuf,

        /// Pretty-print output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,

        /// Rule profile JSON file
        #[arg(long)]
        profile: Option<PathBuf>,
    },

    /// Query the IR with simple predicates
    Query {
        /// Path to the OOXML file
        input: PathBuf,

        /// Node type to match (e.g., Paragraph, Cell, Slide)
        #[arg(long)]
        node_type: Option<String>,

        /// Text search within node content
        #[arg(long)]
        contains: Option<String>,

        /// Document format filter (docx/xlsx/pptx)
        #[arg(long)]
        format: Option<String>,

        /// Require external references (true/false)
        #[arg(long)]
        has_external_refs: Option<bool>,

        /// Require macros (true/false)
        #[arg(long)]
        has_macros: Option<bool>,

        /// Pretty-print output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Select nodes (alias for query)
    Select {
        /// Path to the OOXML file
        input: PathBuf,

        /// Node type to match (e.g., Paragraph, Cell, Slide)
        #[arg(long)]
        node_type: Option<String>,

        /// Text search within node content
        #[arg(long)]
        contains: Option<String>,

        /// Document format filter (docx/xlsx/pptx)
        #[arg(long)]
        format: Option<String>,

        /// Require external references (true/false)
        #[arg(long)]
        has_external_refs: Option<bool>,

        /// Require macros (true/false)
        #[arg(long)]
        has_macros: Option<bool>,

        /// Pretty-print output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Grep-like semantic search (text contains)
    Grep {
        /// Path to the OOXML file
        input: PathBuf,

        /// Pattern to search for
        pattern: String,

        /// Node type to match (e.g., Paragraph, Cell, Slide)
        #[arg(long)]
        node_type: Option<String>,

        /// Document format filter (docx/xlsx/pptx)
        #[arg(long)]
        format: Option<String>,

        /// Pretty-print output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Extract nodes by ID or type
    Extract {
        /// Path to the OOXML file
        input: PathBuf,

        /// Node IDs to extract (repeatable)
        #[arg(long)]
        node_id: Vec<String>,

        /// Node type to extract
        #[arg(long)]
        node_type: Option<String>,

        /// Pretty-print output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },
}

#[derive(Clone, Copy, Debug, ValueEnum)]
pub(crate) enum CoverageExportFormat {
    Json,
    Csv,
}

#[derive(Clone, Copy, Debug, ValueEnum)]
pub(crate) enum CoverageExportMode {
    Full,
    Parts,
}

#[derive(Clone, Copy, ValueEnum)]
pub(crate) enum OutputFormat {
    Json,
    // Future: yaml, binary
}

pub(crate) fn build_parser_config(cli: &Cli) -> ParserConfig {
    let mut config = ParserConfig::default();
    apply_zip_overrides(cli, &mut config);
    apply_odf_overrides(cli, &mut config);
    apply_hwp_overrides(cli, &mut config);
    copy_if_some(cli.max_input_size, &mut config.max_input_size);
    set_if(cli.metrics, &mut config.enable_metrics);
    config
}

fn apply_zip_overrides(cli: &Cli, config: &mut ParserConfig) {
    copy_if_some(
        cli.zip_max_total_size,
        &mut config.zip_config.max_total_size,
    );
    copy_if_some(cli.zip_max_file_size, &mut config.zip_config.max_file_size);
    copy_if_some(
        cli.zip_max_file_count,
        &mut config.zip_config.max_file_count,
    );
    copy_if_some(
        cli.zip_max_compression_ratio,
        &mut config.zip_config.max_compression_ratio,
    );
    copy_if_some(
        cli.zip_max_path_depth,
        &mut config.zip_config.max_path_depth,
    );
}

fn apply_odf_overrides(cli: &Cli, config: &mut ParserConfig) {
    set_if(cli.odf_fast, &mut config.odf.force_fast);
    copy_if_some(
        cli.odf_fast_threshold_bytes,
        &mut config.odf.fast_threshold_bytes,
    );
    copy_if_some(cli.odf_fast_sample_rows, &mut config.odf.fast_sample_rows);
    copy_if_some(cli.odf_fast_sample_cols, &mut config.odf.fast_sample_cols);
    copy_if_some(cli.odf_max_cells.map(non_zero), &mut config.odf.max_cells);
    copy_if_some(cli.odf_max_rows.map(non_zero), &mut config.odf.max_rows);
    copy_if_some(
        cli.odf_max_paragraphs.map(non_zero),
        &mut config.odf.max_paragraphs,
    );
    copy_if_some(cli.odf_max_bytes.map(non_zero), &mut config.odf.max_bytes);
    set_if(cli.odf_parallel_sheets, &mut config.odf.parallel_sheets);
    copy_if_some(
        cli.odf_parallel_max_threads.map(Some),
        &mut config.odf.parallel_max_threads,
    );
    copy_if_some(cli.odf_password.clone().map(Some), &mut config.odf.password);
}

fn apply_hwp_overrides(cli: &Cli, config: &mut ParserConfig) {
    set_if(
        cli.hwp_force_parse_encrypted,
        &mut config.hwp.force_parse_encrypted,
    );
    copy_if_some(cli.hwp_password.clone().map(Some), &mut config.hwp.password);
    set_if(cli.hwp_dump_streams, &mut config.hwp.dump_streams);
}

fn copy_if_some<T>(value: Option<T>, target: &mut T) {
    if let Some(value) = value {
        *target = value;
    }
}

fn set_if(flag: bool, target: &mut bool) {
    if flag {
        *target = true;
    }
}

fn non_zero(value: u64) -> Option<u64> {
    (value != 0).then_some(value)
}
