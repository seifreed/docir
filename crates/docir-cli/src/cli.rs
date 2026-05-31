//! CLI types and configuration wiring.

#[path = "cli_commands.rs"]
mod command_defs;

pub(crate) use command_defs::Commands;

use clap::{Args, Parser, ValueEnum};
use std::path::PathBuf;

/// Common output flags for commands that produce JSON output.
#[derive(Args, Clone)]
pub(crate) struct JsonOutputOpts {
    /// Output as JSON
    #[arg(long)]
    pub(crate) json: bool,

    /// Pretty-print JSON output
    #[arg(long, short)]
    pub(crate) pretty: bool,

    /// Output file (stdout if not specified)
    #[arg(long, short)]
    pub(crate) output: Option<PathBuf>,
}

/// Common output flags for commands that always produce structured output.
#[derive(Args, Clone)]
pub(crate) struct PrettyOutputOpts {
    /// Pretty-print output
    #[arg(long, short)]
    pub(crate) pretty: bool,

    /// Output file (stdout if not specified)
    #[arg(long, short)]
    pub(crate) output: Option<PathBuf>,
}

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

    /// ODF maximum cells to parse (0 = disable, i.e. allow nothing)
    #[arg(long, global = true, value_name = "COUNT")]
    pub(crate) odf_max_cells: Option<u64>,

    /// ODF maximum rows to parse (0 = disable, i.e. allow nothing)
    #[arg(long, global = true, value_name = "COUNT")]
    pub(crate) odf_max_rows: Option<u64>,

    /// ODF maximum paragraphs to parse (0 = disable, i.e. allow nothing)
    #[arg(long, global = true, value_name = "COUNT")]
    pub(crate) odf_max_paragraphs: Option<u64>,

    /// ODF maximum content.xml bytes (0 = disable, i.e. allow nothing)
    #[arg(long, global = true, value_name = "BYTES")]
    pub(crate) odf_max_bytes: Option<u64>,

    /// Enable parallel ODF sheet parsing when possible
    #[arg(long, global = true)]
    pub(crate) odf_parallel_sheets: bool,

    /// Max threads for parallel ODF sheet parsing
    #[arg(long, global = true, value_name = "COUNT")]
    pub(crate) odf_parallel_max_threads: Option<usize>,

    /// Password for decrypting encrypted ODF parts
    /// (prefer DOCIR_ODF_PASSWORD env var to avoid exposing in process list)
    #[arg(
        long,
        global = true,
        value_name = "PASSWORD",
        env = "DOCIR_ODF_PASSWORD"
    )]
    pub(crate) odf_password: Option<String>,

    /// Force parse encrypted HWP streams
    #[arg(long, global = true)]
    pub(crate) hwp_force_parse_encrypted: bool,

    /// Password for decrypting encrypted HWP streams
    /// (prefer DOCIR_HWP_PASSWORD env var to avoid exposing in process list)
    #[arg(
        long,
        global = true,
        value_name = "PASSWORD",
        env = "DOCIR_HWP_PASSWORD"
    )]
    pub(crate) hwp_password: Option<String>,

    /// Dump HWP stream metadata (hash, size, compression)
    #[arg(long, global = true)]
    pub(crate) hwp_dump_streams: bool,

    /// Enable parser timing metrics
    #[arg(long, global = true)]
    pub(crate) metrics: bool,

    /// Disable SHA-256/hash computation during parse and extraction
    #[arg(long, global = true)]
    pub(crate) no_hashes: bool,

    #[command(subcommand)]
    pub(crate) command: Commands,
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
