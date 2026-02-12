//! # docir CLI
//!
//! Command-line interface for the docir document analysis toolkit.

mod commands;

use anyhow::Result;
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
struct Cli {
    /// Maximum total uncompressed size across ZIP entries (bytes)
    #[arg(long, global = true, value_name = "BYTES")]
    zip_max_total_size: Option<u64>,

    /// Maximum size per ZIP entry (bytes)
    #[arg(long, global = true, value_name = "BYTES")]
    zip_max_file_size: Option<u64>,

    /// Maximum number of files in ZIP
    #[arg(long, global = true, value_name = "COUNT")]
    zip_max_file_count: Option<usize>,

    /// Maximum compression ratio for ZIP entries
    #[arg(long, global = true, value_name = "RATIO")]
    zip_max_compression_ratio: Option<f64>,

    /// Maximum path depth inside ZIP
    #[arg(long, global = true, value_name = "DEPTH")]
    zip_max_path_depth: Option<usize>,

    /// Maximum input size for parser entrypoints (bytes)
    #[arg(long, global = true, value_name = "BYTES")]
    max_input_size: Option<u64>,

    /// Force ODF fast mode (skip full cell expansion for large spreadsheets)
    #[arg(long, global = true)]
    odf_fast: bool,

    /// ODF fast-mode threshold for content.xml (bytes)
    #[arg(long, global = true, value_name = "BYTES")]
    odf_fast_threshold_bytes: Option<u64>,

    /// ODF fast-mode sample rows (0 = no sampling)
    #[arg(long, global = true, value_name = "ROWS")]
    odf_fast_sample_rows: Option<u32>,

    /// ODF fast-mode sample columns (0 = no sampling)
    #[arg(long, global = true, value_name = "COLS")]
    odf_fast_sample_cols: Option<u32>,

    /// ODF maximum cells to parse (0 = unlimited)
    #[arg(long, global = true, value_name = "COUNT")]
    odf_max_cells: Option<u64>,

    /// ODF maximum rows to parse (0 = unlimited)
    #[arg(long, global = true, value_name = "COUNT")]
    odf_max_rows: Option<u64>,

    /// ODF maximum paragraphs to parse (0 = unlimited)
    #[arg(long, global = true, value_name = "COUNT")]
    odf_max_paragraphs: Option<u64>,

    /// ODF maximum content.xml bytes (0 = unlimited)
    #[arg(long, global = true, value_name = "BYTES")]
    odf_max_bytes: Option<u64>,

    /// Enable parallel ODF sheet parsing when possible
    #[arg(long, global = true)]
    odf_parallel_sheets: bool,

    /// Max threads for parallel ODF sheet parsing
    #[arg(long, global = true, value_name = "COUNT")]
    odf_parallel_max_threads: Option<usize>,

    /// Password for decrypting encrypted ODF parts
    #[arg(long, global = true, value_name = "PASSWORD")]
    odf_password: Option<String>,

    /// Force parse encrypted HWP streams
    #[arg(long, global = true)]
    hwp_force_parse_encrypted: bool,

    /// Password for decrypting encrypted HWP streams
    #[arg(long, global = true, value_name = "PASSWORD")]
    hwp_password: Option<String>,

    /// Dump HWP stream metadata (hash, size, compression)
    #[arg(long, global = true)]
    hwp_dump_streams: bool,

    /// Enable parser timing metrics
    #[arg(long, global = true)]
    metrics: bool,

    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
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
pub enum CoverageExportFormat {
    Json,
    Csv,
}

#[derive(Clone, Copy, Debug, ValueEnum)]
pub enum CoverageExportMode {
    Full,
    Parts,
}

#[derive(Clone, Copy, ValueEnum)]
enum OutputFormat {
    Json,
    // Future: yaml, binary
}

fn main() -> Result<()> {
    env_logger::Builder::from_env(env_logger::Env::default().default_filter_or("warn")).init();

    let cli = Cli::parse();
    let parser_config = build_parser_config(&cli);

    match cli.command {
        Commands::Parse {
            input,
            format,
            pretty,
            output,
        } => commands::parse::run(input, format, pretty, output, &parser_config),

        Commands::Summary { input } => commands::summary::run(input, &parser_config),
        Commands::Coverage {
            input,
            json,
            details,
            inventory,
            unknown,
            export,
            export_format,
            export_mode,
        } => commands::coverage::run(
            input,
            commands::coverage::CoverageOptions {
                json,
                details,
                inventory,
                unknown,
                export,
                export_format,
                export_mode,
            },
            &parser_config,
        ),

        Commands::Security {
            input,
            json,
            verbose,
        } => commands::security::run(input, json, verbose, &parser_config),

        Commands::DumpNode {
            input,
            node_id,
            format,
        } => commands::dump_node::run(input, &node_id, format, &parser_config),

        Commands::Diff {
            left,
            right,
            pretty,
            output,
        } => commands::diff::run(left, right, pretty, output, &parser_config),

        Commands::Rules {
            input,
            pretty,
            output,
            profile,
        } => commands::rules::run(input, pretty, output, profile, &parser_config),

        Commands::Query {
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            pretty,
            output,
        } => commands::query::run(
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            pretty,
            output,
            &parser_config,
        ),

        Commands::Select {
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            pretty,
            output,
        } => commands::select::run(
            input,
            node_type,
            contains,
            format,
            has_external_refs,
            has_macros,
            pretty,
            output,
            &parser_config,
        ),

        Commands::Grep {
            input,
            pattern,
            node_type,
            format,
            pretty,
            output,
        } => commands::grep::run(
            input,
            pattern,
            node_type,
            format,
            pretty,
            output,
            &parser_config,
        ),

        Commands::Extract {
            input,
            node_id,
            node_type,
            pretty,
            output,
        } => commands::extract::run(input, node_id, node_type, pretty, output, &parser_config),
    }
}

fn build_parser_config(cli: &Cli) -> ParserConfig {
    let mut config = ParserConfig::default();
    if let Some(value) = cli.zip_max_total_size {
        config.zip_config.max_total_size = value;
    }
    if let Some(value) = cli.zip_max_file_size {
        config.zip_config.max_file_size = value;
    }
    if let Some(value) = cli.zip_max_file_count {
        config.zip_config.max_file_count = value;
    }
    if let Some(value) = cli.zip_max_compression_ratio {
        config.zip_config.max_compression_ratio = value;
    }
    if let Some(value) = cli.zip_max_path_depth {
        config.zip_config.max_path_depth = value;
    }
    if let Some(value) = cli.max_input_size {
        config.max_input_size = value;
    }
    if cli.odf_fast {
        config.odf.force_fast = true;
    }
    if let Some(value) = cli.odf_fast_threshold_bytes {
        config.odf.fast_threshold_bytes = value;
    }
    if let Some(value) = cli.odf_fast_sample_rows {
        config.odf.fast_sample_rows = value;
    }
    if let Some(value) = cli.odf_fast_sample_cols {
        config.odf.fast_sample_cols = value;
    }
    if let Some(value) = cli.odf_max_cells {
        config.odf.max_cells = (value != 0).then_some(value);
    }
    if let Some(value) = cli.odf_max_rows {
        config.odf.max_rows = (value != 0).then_some(value);
    }
    if let Some(value) = cli.odf_max_paragraphs {
        config.odf.max_paragraphs = (value != 0).then_some(value);
    }
    if let Some(value) = cli.odf_max_bytes {
        config.odf.max_bytes = (value != 0).then_some(value);
    }
    if cli.odf_parallel_sheets {
        config.odf.parallel_sheets = true;
    }
    if let Some(value) = cli.odf_parallel_max_threads {
        config.odf.parallel_max_threads = Some(value);
    }
    if let Some(password) = cli.odf_password.as_ref() {
        config.odf.password = Some(password.clone());
    }
    if cli.hwp_force_parse_encrypted {
        config.hwp.force_parse_encrypted = true;
    }
    if let Some(password) = cli.hwp_password.as_ref() {
        config.hwp.password = Some(password.clone());
    }
    if cli.hwp_dump_streams {
        config.hwp.dump_streams = true;
    }
    if cli.metrics {
        config.enable_metrics = true;
    }
    config
}
