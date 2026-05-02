//! CLI types and configuration wiring.

use clap::{Parser, Subcommand, ValueEnum};
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

    /// Build an enriched artifact inventory for the document
    Inventory {
        /// Path to the input document
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Probe the real format/container of a file without full parsing
    #[command(name = "probe-format")]
    ProbeFormat {
        /// Path to the input file
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// List CFB storage and stream FILETIMEs
    #[command(name = "list-times")]
    ListTimes {
        /// Path to the input file
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Inspect classic OLE metadata property sets
    #[command(name = "inspect-metadata")]
    InspectMetadata {
        /// Path to the input file
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Inspect low-level BIFF records from a legacy XLS workbook stream
    #[command(name = "inspect-sheet-records")]
    InspectSheetRecords {
        /// Path to the input file
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Inspect low-level binary records from a legacy PPT presentation stream
    #[command(name = "inspect-slide-records")]
    InspectSlideRecords {
        /// Path to the input file
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Inspect normal CFB directory entries and their structural metadata
    #[command(name = "inspect-directory")]
    InspectDirectory {
        /// Path to the input file
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Inspect CFB sector allocation and stream chains
    #[command(name = "inspect-sectors")]
    InspectSectors {
        /// Path to the input file
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Build an analyst-facing indicator scorecard for the document
    #[command(name = "report-indicators")]
    ReportIndicators {
        /// Path to the input document
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Extract DDE-style active links into a dedicated report
    #[command(name = "extract-links")]
    ExtractLinks {
        /// Path to the input document
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Detect and optionally export embedded SWF/Flash payloads
    #[command(name = "extract-flash")]
    ExtractFlash {
        /// Path to the input document
        input: PathBuf,

        /// Output directory for extracted SWF payloads
        #[arg(long)]
        out: Option<PathBuf>,

        /// Allow writing into an existing output directory
        #[arg(long)]
        overwrite: bool,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Emit the canonical Phase 0 artifact manifest as JSON
    Manifest {
        /// Path to the input document
        input: PathBuf,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Dump low-level container entries for OOXML, CFB, or RTF inputs
    #[command(name = "dump-container")]
    DumpContainer {
        /// Path to the input document
        input: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Recognize VBA projects and modules without AST or deobfuscation
    #[command(name = "recognize-vba")]
    RecognizeVba {
        /// Path to the input document
        input: PathBuf,

        /// Include normalized module source in JSON/text output
        #[arg(long)]
        include_source: bool,

        /// Output as JSON
        #[arg(long)]
        json: bool,

        /// Pretty-print JSON output
        #[arg(long, short)]
        pretty: bool,

        /// Output file (stdout if not specified)
        #[arg(long, short)]
        output: Option<PathBuf>,
    },

    /// Extract VBA modules to disk and emit a manifest
    #[command(name = "extract-vba")]
    ExtractVba {
        /// Path to the input document
        input: PathBuf,

        /// Output directory
        #[arg(long)]
        out: PathBuf,

        /// Allow writing into an existing output directory
        #[arg(long)]
        overwrite: bool,

        /// Keep partial bundles when some modules cannot be extracted
        #[arg(long)]
        best_effort: bool,
    },

    /// Extract embedded OOXML/RTF artifacts to disk and emit a manifest
    #[command(name = "extract-artifacts")]
    ExtractArtifacts {
        /// Path to the input document
        input: PathBuf,

        /// Output directory
        #[arg(long)]
        out: PathBuf,

        /// Allow writing into an existing output directory
        #[arg(long)]
        overwrite: bool,

        /// Also dump raw OOXML embedding and ActiveX container binaries
        #[arg(long)]
        with_raw: bool,

        /// Exclude regular OOXML media assets such as images/audio/video
        #[arg(long)]
        no_media: bool,

        /// Restrict output to OLE-backed artifacts
        #[arg(long)]
        only_ole: bool,

        /// Restrict extraction to RTF objdata blobs
        #[arg(long)]
        only_rtf_objects: bool,
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
