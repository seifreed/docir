//! Parser configuration types.

use crate::zip_handler::ZipConfig;

/// ODF parsing configuration.
#[derive(Debug, Clone)]
pub struct OdfConfig {
    /// ODF fast-mode threshold (content.xml size in bytes).
    pub fast_threshold_bytes: u64,
    /// Force ODF fast mode regardless of size.
    pub force_fast: bool,
    /// ODF fast-mode sample rows (0 = no sampling).
    pub fast_sample_rows: u32,
    /// ODF fast-mode sample columns (0 = no sampling).
    pub fast_sample_cols: u32,
    /// ODF maximum cells (None = unlimited).
    pub max_cells: Option<u64>,
    /// ODF maximum rows (None = unlimited).
    pub max_rows: Option<u64>,
    /// ODF maximum paragraphs (None = unlimited).
    pub max_paragraphs: Option<u64>,
    /// ODF maximum content.xml bytes (None = unlimited).
    pub max_bytes: Option<u64>,
    /// Enable parallel parsing of ODF sheets when possible.
    pub parallel_sheets: bool,
    /// Max threads for parallel ODF sheet parsing (None = auto).
    pub parallel_max_threads: Option<usize>,
    /// Password for decrypting encrypted ODF parts.
    pub password: Option<String>,
}

impl Default for OdfConfig {
    fn default() -> Self {
        Self {
            fast_threshold_bytes: 50 * 1024 * 1024,
            force_fast: false,
            fast_sample_rows: 0,
            fast_sample_cols: 0,
            max_cells: None,
            max_rows: None,
            max_paragraphs: None,
            max_bytes: None,
            parallel_sheets: false,
            parallel_max_threads: None,
            password: None,
        }
    }
}

/// RTF parsing configuration.
#[derive(Debug, Clone)]
pub struct RtfConfig {
    /// RTF max group depth (0 = unlimited).
    pub max_group_depth: usize,
    /// RTF max object hex length (0 = unlimited).
    pub max_object_hex_len: usize,
}

impl Default for RtfConfig {
    fn default() -> Self {
        Self {
            max_group_depth: 256,
            max_object_hex_len: 64 * 1024 * 1024,
        }
    }
}

/// HWP parsing configuration.
#[derive(Debug, Clone, Default)]
pub struct HwpConfig {
    /// Force parse encrypted HWP streams.
    pub force_parse_encrypted: bool,
    /// Password for decrypting encrypted HWP streams.
    pub password: Option<String>,
    /// Dump HWP stream metadata (hash, size, compression).
    pub dump_streams: bool,
}

/// Parser configuration.
#[derive(Debug, Clone)]
pub struct ParserConfig {
    /// ZIP extraction limits.
    pub zip_config: ZipConfig,
    /// Maximum XML recursion depth.
    pub max_xml_depth: usize,
    /// Extract VBA macro source code.
    pub extract_macro_source: bool,
    /// Compute hashes for OLE objects.
    pub compute_hashes: bool,
    /// Enforce minimal required parts per format.
    pub enforce_required_parts: bool,
    /// Perform security scanning during parse.
    pub scan_security_on_parse: bool,
    /// Collect parse timing metrics.
    pub enable_metrics: bool,
    /// Maximum allowed input size for top-level parser entrypoints.
    pub max_input_size: u64,
    /// ODF-specific configuration.
    pub odf: OdfConfig,
    /// RTF-specific configuration.
    pub rtf: RtfConfig,
    /// HWP-specific configuration.
    pub hwp: HwpConfig,
}

impl Default for ParserConfig {
    fn default() -> Self {
        Self {
            zip_config: ZipConfig::default(),
            max_xml_depth: 100,
            extract_macro_source: false,
            compute_hashes: true,
            enforce_required_parts: true,
            scan_security_on_parse: true,
            enable_metrics: false,
            max_input_size: 512 * 1024 * 1024,
            odf: OdfConfig::default(),
            rtf: RtfConfig::default(),
            hwp: HwpConfig::default(),
        }
    }
}

/// Basic parse metrics (timings in milliseconds).
#[derive(Debug, Clone, Default)]
pub struct ParseMetrics {
    pub content_types_ms: u128,
    pub relationships_ms: u128,
    pub main_parse_ms: u128,
    pub shared_parts_ms: u128,
    pub security_scan_ms: u128,
    pub extension_parts_ms: u128,
    pub normalization_ms: u128,
}
