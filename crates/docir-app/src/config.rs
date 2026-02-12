use docir_parser::parser::ParseMetrics as ParserParseMetrics;

#[derive(Debug, Clone)]
pub struct ZipConfig {
    pub max_total_size: u64,
    pub max_file_size: u64,
    pub max_file_count: usize,
    pub max_compression_ratio: f64,
    pub max_path_depth: usize,
}

impl Default for ZipConfig {
    fn default() -> Self {
        Self {
            max_total_size: 100 * 1024 * 1024,
            max_file_size: 50 * 1024 * 1024,
            max_file_count: 10000,
            max_compression_ratio: 100.0,
            max_path_depth: 20,
        }
    }
}

#[derive(Debug, Clone)]
pub struct OdfConfig {
    pub fast_threshold_bytes: u64,
    pub force_fast: bool,
    pub fast_sample_rows: u32,
    pub fast_sample_cols: u32,
    pub max_cells: Option<u64>,
    pub max_rows: Option<u64>,
    pub max_paragraphs: Option<u64>,
    pub max_bytes: Option<u64>,
    pub parallel_sheets: bool,
    pub parallel_max_threads: Option<usize>,
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

#[derive(Debug, Clone)]
pub struct RtfConfig {
    pub max_group_depth: usize,
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

#[derive(Debug, Clone)]
pub struct HwpConfig {
    pub force_parse_encrypted: bool,
    pub password: Option<String>,
    pub dump_streams: bool,
}

impl Default for HwpConfig {
    fn default() -> Self {
        Self {
            force_parse_encrypted: false,
            password: None,
            dump_streams: false,
        }
    }
}

#[derive(Debug, Clone)]
pub struct ParserConfig {
    pub zip_config: ZipConfig,
    pub max_xml_depth: usize,
    pub extract_macro_source: bool,
    pub compute_hashes: bool,
    pub enforce_required_parts: bool,
    pub scan_security_on_parse: bool,
    pub enable_metrics: bool,
    pub max_input_size: u64,
    pub odf: OdfConfig,
    pub rtf: RtfConfig,
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

impl ParseMetrics {
    pub(crate) fn from_parser(metrics: &ParserParseMetrics) -> Self {
        Self {
            content_types_ms: metrics.content_types_ms,
            relationships_ms: metrics.relationships_ms,
            main_parse_ms: metrics.main_parse_ms,
            shared_parts_ms: metrics.shared_parts_ms,
            security_scan_ms: metrics.security_scan_ms,
            extension_parts_ms: metrics.extension_parts_ms,
            normalization_ms: metrics.normalization_ms,
        }
    }
}

impl ParserConfig {
    pub(crate) fn to_parser_config(&self) -> docir_parser::ParserConfig {
        docir_parser::ParserConfig {
            zip_config: docir_parser::zip_handler::ZipConfig {
                max_total_size: self.zip_config.max_total_size,
                max_file_size: self.zip_config.max_file_size,
                max_file_count: self.zip_config.max_file_count,
                max_compression_ratio: self.zip_config.max_compression_ratio,
                max_path_depth: self.zip_config.max_path_depth,
            },
            max_xml_depth: self.max_xml_depth,
            extract_macro_source: self.extract_macro_source,
            compute_hashes: self.compute_hashes,
            enforce_required_parts: self.enforce_required_parts,
            scan_security_on_parse: self.scan_security_on_parse,
            enable_metrics: self.enable_metrics,
            max_input_size: self.max_input_size,
            odf: docir_parser::config::OdfConfig {
                fast_threshold_bytes: self.odf.fast_threshold_bytes,
                force_fast: self.odf.force_fast,
                fast_sample_rows: self.odf.fast_sample_rows,
                fast_sample_cols: self.odf.fast_sample_cols,
                max_cells: self.odf.max_cells,
                max_rows: self.odf.max_rows,
                max_paragraphs: self.odf.max_paragraphs,
                max_bytes: self.odf.max_bytes,
                parallel_sheets: self.odf.parallel_sheets,
                parallel_max_threads: self.odf.parallel_max_threads,
                password: self.odf.password.clone(),
            },
            rtf: docir_parser::config::RtfConfig {
                max_group_depth: self.rtf.max_group_depth,
                max_object_hex_len: self.rtf.max_object_hex_len,
            },
            hwp: docir_parser::config::HwpConfig {
                force_parse_encrypted: self.hwp.force_parse_encrypted,
                password: self.hwp.password.clone(),
                dump_streams: self.hwp.dump_streams,
            },
        }
    }
}
