use super::{
    CoverageExportFormat, CoverageExportMode, JsonOutputOpts, OutputFormat, PrettyOutputOpts,
};
use clap::Subcommand;
use std::path::PathBuf;

#[derive(Subcommand)]
pub(crate) enum Commands {
    /// Parse a document and output its IR
    Parse {
        /// Path to the OOXML file
        input: PathBuf,

        /// Output format
        #[arg(long, short, default_value = "json")]
        format: OutputFormat,

        #[command(flatten)]
        output_opts: PrettyOutputOpts,
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

        #[command(flatten)]
        output_opts: JsonOutputOpts,
    },

    /// Probe the real format/container of a file without full parsing
    #[command(name = "probe-format")]
    ProbeFormat {
        /// Path to the input file
        input: PathBuf,

        #[command(flatten)]
        output_opts: JsonOutputOpts,
    },

    /// List CFB storage and stream FILETIMEs
    #[command(name = "list-times")]
    ListTimes {
        /// Path to the input file
        input: PathBuf,

        #[command(flatten)]
        output_opts: JsonOutputOpts,
    },

    /// Inspect classic OLE metadata property sets
    #[command(name = "inspect-metadata")]
    InspectMetadata {
        /// Path to the input file
        input: PathBuf,

        #[command(flatten)]
        output_opts: JsonOutputOpts,
    },

    /// Inspect low-level BIFF records from a legacy XLS workbook stream
    #[command(name = "inspect-sheet-records")]
    InspectSheetRecords {
        /// Path to the input file
        input: PathBuf,

        #[command(flatten)]
        output_opts: JsonOutputOpts,
    },

    /// Inspect low-level binary records from a legacy PPT presentation stream
    #[command(name = "inspect-slide-records")]
    InspectSlideRecords {
        /// Path to the input file
        input: PathBuf,

        #[command(flatten)]
        output_opts: JsonOutputOpts,
    },

    /// Inspect normal CFB directory entries and their structural metadata
    #[command(name = "inspect-directory")]
    InspectDirectory {
        /// Path to the input file
        input: PathBuf,

        #[command(flatten)]
        output_opts: JsonOutputOpts,
    },

    /// Inspect CFB sector allocation and stream chains
    #[command(name = "inspect-sectors")]
    InspectSectors {
        /// Path to the input file
        input: PathBuf,

        #[command(flatten)]
        output_opts: JsonOutputOpts,
    },

    /// Build an analyst-facing indicator scorecard for the document
    #[command(name = "report-indicators")]
    ReportIndicators {
        /// Path to the input document
        input: PathBuf,

        #[command(flatten)]
        output_opts: JsonOutputOpts,
    },

    /// Extract DDE-style active links into a dedicated report
    #[command(name = "extract-links")]
    ExtractLinks {
        /// Path to the input document
        input: PathBuf,

        #[command(flatten)]
        output_opts: JsonOutputOpts,
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

        #[command(flatten)]
        output_opts: JsonOutputOpts,
    },

    /// Emit the canonical Phase 0 artifact manifest as JSON
    Manifest {
        /// Path to the input document
        input: PathBuf,

        #[command(flatten)]
        output_opts: PrettyOutputOpts,
    },

    /// Dump low-level container entries for OOXML, CFB, or RTF inputs
    #[command(name = "dump-container")]
    DumpContainer {
        /// Path to the input document
        input: PathBuf,

        #[command(flatten)]
        output_opts: JsonOutputOpts,
    },

    /// Recognize VBA projects and modules without AST or deobfuscation
    #[command(name = "recognize-vba")]
    RecognizeVba {
        /// Path to the input document
        input: PathBuf,

        /// Include normalized module source in JSON/text output
        #[arg(long)]
        include_source: bool,

        #[command(flatten)]
        output_opts: JsonOutputOpts,
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

        #[command(flatten)]
        output_opts: PrettyOutputOpts,
    },

    /// Run rule engine on a document
    Rules {
        /// Path to the OOXML file
        input: PathBuf,

        #[command(flatten)]
        output_opts: PrettyOutputOpts,

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

        #[command(flatten)]
        output_opts: PrettyOutputOpts,
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

        #[command(flatten)]
        output_opts: PrettyOutputOpts,
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

        #[command(flatten)]
        output_opts: PrettyOutputOpts,
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

        #[command(flatten)]
        output_opts: PrettyOutputOpts,
    },
}
