//! Port traits for hexagonal architecture.
//!
//! These traits define the application boundaries. Adapters provide concrete
//! implementations; the [`DocirApp`](super::DocirApp) facade depends on these
//! ports, not on infrastructure details.

use crate::{AppResult, ParsedDocument};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;
use docir_rules::RuleProfile;
use docir_rules::RuleReport;
use docir_security::analyzer::AnalysisResult;
use std::io::{Read, Seek};
use std::path::Path;

/// Parser port for application workflows.
///
/// This trait is intended for adapters so callers can plug alternate parser
/// implementations without touching application orchestration code.
pub trait ParserPort {
    /// Parse a file path into a canonical `ParsedDocument`.
    fn parse_file<P: AsRef<Path>>(&self, path: P) -> AppResult<ParsedDocument>;
    /// Parse a byte slice into a canonical `ParsedDocument`.
    fn parse_bytes(&self, data: &[u8]) -> AppResult<ParsedDocument>;
    /// Parse any readable + seekable source into a canonical `ParsedDocument`.
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> AppResult<ParsedDocument>;
    /// Parse and capture the original bytes from a file path.
    fn parse_file_with_bytes<P: AsRef<Path>>(
        &self,
        path: P,
    ) -> AppResult<(ParsedDocument, Vec<u8>)>;
    /// Parse and capture the original bytes from a reader.
    fn parse_reader_with_bytes<R: Read + Seek>(
        &self,
        reader: R,
    ) -> AppResult<(ParsedDocument, Vec<u8>)>;
}

/// Security scanning port for application workflows.
///
/// Implementations can scan byte streams and apply side effects on the shared
/// IR store.
pub trait SecurityScannerPort {
    /// Run byte-level security scanning for the given container bytes.
    fn scan_security_bytes(&self, data: &[u8], store: &mut IrStore) -> AppResult<()>;
}

pub(crate) trait CfbStreamReaderPort {
    fn read_streams(&self, data: &[u8], stream_names: &[&str])
        -> AppResult<Vec<(String, Vec<u8>)>>;
}

/// Security analysis port for application workflows.
///
/// Implementations compute security conclusions from an existing IR.
pub trait SecurityAnalyzerPort {
    /// Analyze security signals for a parsed document root.
    fn analyze(&mut self, store: &IrStore, root_id: NodeId) -> AnalysisResult;
}

/// Security enrichment port for application workflows.
pub trait SecurityEnricherPort {
    /// Enrich security annotations into the provided IR.
    fn enrich(&self, store: &mut IrStore, root_id: NodeId);
}

/// Rules engine port for application workflows.
///
/// Rules are provided with read-only IR access and produce a rules report.
pub trait RulesEnginePort {
    /// Run rule profile against a parsed document root.
    fn run_with_profile(
        &self,
        store: &IrStore,
        root_id: NodeId,
        profile: &RuleProfile,
    ) -> RuleReport;
}

/// Serialization port for application workflows.
///
/// Adapters serialize parsed documents according to output media type.
pub trait SerializerPort {
    /// Serialize a parsed document to JSON output.
    fn to_json(&self, parsed: &ParsedDocument, pretty: bool) -> AppResult<String>;
}

/// Summary presentation port for output adapters.
///
/// Implementations produce text outputs from summary models.
pub trait SummaryPresenterPort {
    /// Format a summary for CLI/report display.
    fn format_summary(
        &self,
        summary: &crate::summary::DocumentSummary,
        source: Option<&str>,
    ) -> String;
}

impl SecurityAnalyzerPort for docir_security::SecurityAnalyzer {
    fn analyze(&mut self, store: &IrStore, root_id: NodeId) -> AnalysisResult {
        self.analyze(store, root_id)
    }
}
