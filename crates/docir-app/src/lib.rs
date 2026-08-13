//! Application-level workflows for docir.

use docir_core::ir::Document;
use docir_core::security::SecurityInfo;
use docir_core::types::{DocumentFormat, NodeId};
use docir_core::visitor::IrStore;
use docir_diff::{DiffError, DiffResult};
use docir_parser::ParseError as ParserParseError;
use docir_parser::parser::ParsedDocument as ParserParsedDocument;
pub use docir_rules::RuleProfile;
pub use docir_rules::RuleReport;
use docir_serialization::SerializationError;
use std::io::{Read, Seek};
use std::path::Path;
use thiserror::Error;

mod adapters;
mod artifacts;
mod bucket_count;
mod config;
mod container;
mod export;
mod extract_flash;
mod extract_links;
mod inspect_directory;
mod inspect_sectors;
mod inspect_sheet_records;
mod inspect_slide_records;
mod inventory;
mod io_support;
mod list_times;
mod metadata;
mod ports;
mod probe;
mod report_indicators;
mod severity;
mod summary;
#[cfg(any(test, feature = "test-support"))]
pub mod test_support;
mod use_cases;
mod vba;

/// Primary facade adapter for parser implementations.
pub use adapters::AppParser;
pub use artifacts::{
    ArtifactExtractionBundle, ArtifactExtractionOptions, ExtractedPayload,
    extract_artifacts_from_bytes,
};
pub use bucket_count::BucketCount;
/// Parser-related CLI configuration bundle.
pub use config::{HwpConfig, OdfConfig, ParseMetrics, ParserConfig, RtfConfig, ZipConfig};
pub use container::{ContainerDump, ContainerEntry, ContainerEntryKind};
/// Result type produced by static security analyzers.
pub use docir_security::analyzer::AnalysisResult;
pub use export::{
    ExportDocumentRef, Phase0Artifact, Phase0ArtifactLocator, Phase0ArtifactManifestExport,
    Phase0Diagnostic, Phase0VbaBody, Phase0VbaExport, Phase0VbaModule, Phase0VbaProject,
    PhaseCapabilities,
};
pub use extract_flash::{
    FlashExtractionReport, FlashObject, extract_flash_bytes, extract_flash_path,
};
pub use extract_links::{LinkArtifact, LinkExtractionReport};
pub use inspect_directory::{
    DirectoryAnomalySeverity, DirectoryEntry, DirectoryInspection, inspect_directory_bytes,
    inspect_directory_path,
};
pub use inspect_sectors::{
    ChainHealthCount, ChainStep, RoleCount, SectorAnomaly, SectorInspection, SectorOverviewEntry,
    SectorOwnerRef, SharedChainOverlap, SharedSectorClaim, StartSectorReuse, StreamSectorMap,
    StructuralIncoherenceCount, TruncatedChainCount, inspect_sectors_bytes, inspect_sectors_path,
};
pub use inspect_sheet_records::{
    SheetRecordAnomaly, SheetRecordCount, SheetRecordEntry, SheetRecordInspection,
    inspect_sheet_records_bytes, inspect_sheet_records_path,
};
pub use inspect_slide_records::{
    SlideRecordAnomaly, SlideRecordCount, SlideRecordEntry, SlideRecordInspection,
    inspect_slide_records_bytes, inspect_slide_records_path,
};
pub use inventory::{ArtifactInventory, ContainerKind, InventoryArtifact, InventoryArtifactKind};
pub use list_times::{TimeEntry, TimeListing, list_times_bytes, list_times_path};
pub use metadata::{
    MetadataInspection, MetadataProperty, MetadataSection, inspect_metadata_bytes,
    inspect_metadata_path,
};
pub use probe::{FormatProbe, probe_format_bytes, probe_format_path};
pub use report_indicators::{DocumentIndicator, IndicatorReport};
/// Structured summary models for CLI and report outputs.
pub use summary::{
    DocumentSummary, MetadataSummary, NodeCount, ParseMetricsSummary, SecuritySummary,
    TextStatsSummary, ThreatIndicatorSummary, summarize_document,
};
pub use vba::{VbaModuleReport, VbaProjectReport, VbaRecognitionReport, VbaRecognitionStatus};

pub use ports::{
    ParserPort, RulesEnginePort, SecurityAnalyzerPort, SecurityEnricherPort, SecurityScannerPort,
    SerializerPort, SummaryPresenterPort,
};

use use_cases::{
    AnalyzeSecurity, AnalyzeSecurityUseCase, DiffDocuments, ParseDocument, ParseDocumentUseCase,
    RunRules, SummarizeUseCase,
};

/// Result alias for all docir-app operations.
pub type AppResult<T> = Result<T, AppError>;

/// High-level error type for application workflows.
#[derive(Debug, Error)]
pub enum AppError {
    /// IR traversal failure while building an application result.
    #[error(transparent)]
    Core(#[from] docir_core::CoreError),
    /// Parser pipeline failure.
    #[error(transparent)]
    Parse(#[from] ParserParseError),
    /// Diff pipeline failure.
    #[error(transparent)]
    Diff(#[from] DiffError),
    /// Serialization failure.
    #[error(transparent)]
    Serialization(#[from] SerializationError),
}

/// Application-level parsed document wrapper.
#[derive(Debug)]
pub struct ParsedDocument {
    inner: ParserParsedDocument,
    metrics: Option<ParseMetrics>,
}

impl ParsedDocument {
    pub(crate) fn new(inner: ParserParsedDocument) -> Self {
        let metrics = inner.metrics.clone();
        Self { inner, metrics }
    }

    /// Returns the root node id of the parsed document.
    pub fn root_id(&self) -> NodeId {
        self.inner.root_id
    }

    /// Returns the document format used by the parser.
    pub fn format(&self) -> DocumentFormat {
        self.inner.format
    }

    /// Returns a shared reference to the document store.
    pub fn store(&self) -> &IrStore {
        &self.inner.store
    }

    /// Returns a mutable reference to the document store.
    pub fn store_mut(&mut self) -> &mut IrStore {
        &mut self.inner.store
    }

    /// Returns the high-level document model if present.
    pub fn document(&self) -> Option<&Document> {
        self.inner.document()
    }

    /// Returns the security scan summary if present.
    pub fn security_info(&self) -> Option<&SecurityInfo> {
        self.inner.security_info()
    }

    /// Returns collected parser metrics if present.
    pub fn metrics(&self) -> Option<&ParseMetrics> {
        self.metrics.as_ref()
    }
}

/// Application facade for docir workflows.
pub struct DocirApp<P: ParserPort + SecurityScannerPort = AppParser> {
    parser: P,
    security_analyzer_factory: Box<dyn Fn() -> Box<dyn SecurityAnalyzerPort>>,
    security_enricher: Box<dyn SecurityEnricherPort>,
    rules_engine_factory: Box<dyn Fn() -> Box<dyn RulesEnginePort>>,
    serializer: Box<dyn SerializerPort>,
    summary_presenter: Box<dyn SummaryPresenterPort>,
}

impl DocirApp<AppParser> {
    /// Creates a new app instance with the provided parser config.
    pub fn new(config: ParserConfig) -> Self {
        let mut config = config;
        config.scan_security_on_parse = false;
        Self::with_parser(AppParser::with_config(config))
    }

    /// Builds a low-level dump of the underlying source container.
    pub fn build_container_dump(
        &self,
        parsed: &ParsedDocument,
        input_bytes: &[u8],
    ) -> AppResult<ContainerDump> {
        ContainerDump::from_parsed_bytes(parsed, input_bytes, self.parser.zip_config())
    }
}

impl<P: ParserPort + SecurityScannerPort> DocirApp<P> {
    fn parse_use_case(&self) -> ParseDocumentUseCase<'_, P, P> {
        ParseDocument::new(&self.parser, &self.parser, self.security_enricher.as_ref())
    }

    fn summarize_use_case(&self) -> SummarizeUseCase {
        SummarizeUseCase
    }

    fn analyze_security_use_case(
        &self,
    ) -> AnalyzeSecurityUseCase<&'_ dyn Fn() -> Box<dyn SecurityAnalyzerPort>> {
        AnalyzeSecurity::new(self.security_analyzer_factory.as_ref())
    }

    /// Creates a new app instance with a custom parser implementation.
    pub fn with_parser(parser: P) -> Self {
        Self::with_parser_and_ports(
            parser,
            adapters::default_security_analyzer_factory(),
            adapters::default_rules_engine_factory(),
            adapters::default_json_serializer(),
            adapters::default_summary_presenter(),
        )
    }

    /// Creates a new app instance with custom parser and security analyzer factory.
    pub fn with_parser_and_security<F>(parser: P, security_analyzer_factory: F) -> Self
    where
        F: Fn() -> Box<dyn SecurityAnalyzerPort> + 'static,
    {
        Self::with_parser_and_ports(
            parser,
            security_analyzer_factory,
            adapters::default_rules_engine_factory(),
            adapters::default_json_serializer(),
            adapters::default_summary_presenter(),
        )
    }

    /// Creates a new app instance with custom ports.
    pub fn with_parser_and_ports<F, R>(
        parser: P,
        security_analyzer_factory: F,
        rules_engine_factory: R,
        serializer: Box<dyn SerializerPort>,
        summary_presenter: Box<dyn SummaryPresenterPort>,
    ) -> Self
    where
        F: Fn() -> Box<dyn SecurityAnalyzerPort> + 'static,
        R: Fn() -> Box<dyn RulesEnginePort> + 'static,
    {
        Self::with_parser_and_ports_and_enricher(
            parser,
            security_analyzer_factory,
            adapters::default_security_enricher(),
            rules_engine_factory,
            serializer,
            summary_presenter,
        )
    }

    /// Creates a new app instance with custom ports and security enricher.
    pub fn with_parser_and_ports_and_enricher<F, R>(
        parser: P,
        security_analyzer_factory: F,
        security_enricher: Box<dyn SecurityEnricherPort>,
        rules_engine_factory: R,
        serializer: Box<dyn SerializerPort>,
        summary_presenter: Box<dyn SummaryPresenterPort>,
    ) -> Self
    where
        F: Fn() -> Box<dyn SecurityAnalyzerPort> + 'static,
        R: Fn() -> Box<dyn RulesEnginePort> + 'static,
    {
        Self {
            parser,
            security_analyzer_factory: Box::new(security_analyzer_factory),
            security_enricher,
            rules_engine_factory: Box::new(rules_engine_factory),
            serializer,
            summary_presenter,
        }
    }

    /// Parses a file from disk.
    pub fn parse_file<Pth: AsRef<Path>>(&self, path: Pth) -> AppResult<ParsedDocument> {
        self.parse_use_case().parse_file(path)
    }

    /// Parses a file from disk and returns the original input bytes.
    pub fn parse_file_with_bytes<Pth: AsRef<Path>>(
        &self,
        path: Pth,
    ) -> AppResult<(ParsedDocument, Vec<u8>)> {
        self.parse_use_case().parse_file_with_bytes(path)
    }

    /// Parses from bytes.
    pub fn parse_bytes(&self, data: &[u8]) -> AppResult<ParsedDocument> {
        self.parse_use_case().parse_bytes(data)
    }

    /// Parses from a reader.
    pub fn parse_reader<R: Read + Seek>(&self, reader: R) -> AppResult<ParsedDocument> {
        self.parse_use_case().parse_reader(reader)
    }

    /// Parses from a reader and returns the original input bytes.
    pub fn parse_reader_with_bytes<R: Read + Seek>(
        &self,
        reader: R,
    ) -> AppResult<(ParsedDocument, Vec<u8>)> {
        self.parse_use_case().parse_reader_with_bytes(reader)
    }

    /// Serializes a parsed document to JSON.
    pub fn serialize_json(&self, parsed: &ParsedDocument, pretty: bool) -> AppResult<String> {
        self.serializer.to_json(parsed, pretty)
    }

    /// Builds a structured summary for a parsed document.
    pub fn build_summary(&self, parsed: &ParsedDocument) -> AppResult<Option<DocumentSummary>> {
        self.summarize_use_case().run(parsed)
    }

    /// Builds a structured artifact inventory for a parsed document.
    pub fn build_inventory(&self, parsed: &ParsedDocument) -> ArtifactInventory {
        ArtifactInventory::from_parsed(parsed)
    }

    /// Builds a structured artifact inventory enriched with low-level container metadata.
    pub fn build_inventory_with_bytes(
        &self,
        parsed: &ParsedDocument,
        input_bytes: &[u8],
    ) -> ArtifactInventory {
        ArtifactInventory::from_parsed_with_bytes(parsed, input_bytes)
    }

    /// Builds a structured VBA recognition report for a parsed document.
    pub fn build_vba_recognition(
        &self,
        parsed: &ParsedDocument,
        include_source: bool,
    ) -> VbaRecognitionReport {
        VbaRecognitionReport::from_parsed(parsed, include_source)
    }

    /// Builds an analyst-facing indicator scorecard for a parsed document.
    pub fn build_indicator_report(&self, parsed: &ParsedDocument) -> IndicatorReport {
        IndicatorReport::from_parsed(parsed)
    }

    /// Builds a low-level legacy XLS BIFF record inspection report from raw bytes.
    pub fn inspect_sheet_records_from_bytes(
        &self,
        source_bytes: &[u8],
    ) -> AppResult<SheetRecordInspection> {
        inspect_sheet_records_bytes(source_bytes)
    }

    /// Builds a low-level legacy PPT record inspection report from raw bytes.
    pub fn inspect_slide_records_from_bytes(
        &self,
        source_bytes: &[u8],
    ) -> AppResult<SlideRecordInspection> {
        inspect_slide_records_bytes(source_bytes)
    }

    pub fn build_indicator_report_with_bytes(
        &self,
        parsed: &ParsedDocument,
        source_bytes: &[u8],
    ) -> IndicatorReport {
        IndicatorReport::from_parsed_with_bytes(parsed, Some(source_bytes))
    }

    /// Builds a dedicated report for link-like active content such as DDE.
    pub fn build_link_extraction_report(&self, parsed: &ParsedDocument) -> LinkExtractionReport {
        LinkExtractionReport::from_parsed(parsed)
    }

    /// Builds and formats a structured summary for output adapters.
    pub fn format_summary(
        &self,
        parsed: &ParsedDocument,
        source: Option<&str>,
    ) -> AppResult<Option<String>> {
        self.build_summary(parsed).map(|summary| {
            summary.map(|summary| self.summary_presenter.format_summary(&summary, source))
        })
    }

    /// Runs security analysis for a parsed document.
    pub fn analyze_security(&self, parsed: &ParsedDocument) -> AnalysisResult {
        self.analyze_security_use_case()
            .run(parsed.store(), parsed.root_id())
    }

    /// Runs rules for a parsed document.
    pub fn run_rules(&self, parsed: &ParsedDocument, profile: &RuleProfile) -> RuleReport {
        RunRules::new(&self.rules_engine_factory).run(parsed.store(), parsed.root_id(), profile)
    }

    /// Computes a diff between two parsed documents.
    pub fn diff(&self, left: &ParsedDocument, right: &ParsedDocument) -> AppResult<DiffResult> {
        DiffDocuments::diff(left, right)
    }
}

#[cfg(test)]
mod tests;
