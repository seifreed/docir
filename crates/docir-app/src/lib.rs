//! Application-level workflows for docir.

use docir_core::ir::Document;
use docir_core::security::SecurityInfo;
use docir_core::types::{DocumentFormat, NodeId};
use docir_core::visitor::IrStore;
use docir_diff::{DiffError, DiffResult};
use docir_parser::parser::ParsedDocument as ParserParsedDocument;
use docir_parser::ParseError as ParserParseError;
pub use docir_rules::RuleProfile;
use docir_rules::RuleReport;
use docir_security::SecurityAnalyzer;
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
mod list_times;
mod metadata;
mod probe;
mod report_indicators;
mod severity;
mod summary;
pub mod test_support;
mod use_cases;
mod vba;

/// Primary facade adapter for parser implementations.
pub use adapters::AppParser;
pub use artifacts::{
    extract_artifacts_from_bytes, ArtifactExtractionBundle, ArtifactExtractionOptions,
    ExtractedPayload,
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
    extract_flash_bytes, extract_flash_path, FlashExtractionReport, FlashObject,
};
pub use extract_links::{LinkArtifact, LinkExtractionReport};
pub use inspect_directory::{
    inspect_directory_bytes, inspect_directory_path, DirectoryAnomalySeverity, DirectoryEntry,
    DirectoryInspection,
};
pub use inspect_sectors::{
    inspect_sectors_bytes, inspect_sectors_path, ChainHealthCount, ChainStep, RoleCount,
    SectorAnomaly, SectorInspection, SectorOverviewEntry, SectorOwnerRef, SharedChainOverlap,
    SharedSectorClaim, StartSectorReuse, StreamSectorMap, StructuralIncoherenceCount,
    TruncatedChainCount,
};
pub use inspect_sheet_records::{
    inspect_sheet_records_bytes, inspect_sheet_records_path, SheetRecordAnomaly, SheetRecordCount,
    SheetRecordEntry, SheetRecordInspection,
};
pub use inspect_slide_records::{
    inspect_slide_records_bytes, inspect_slide_records_path, SlideRecordAnomaly, SlideRecordCount,
    SlideRecordEntry, SlideRecordInspection,
};
pub use inventory::{ArtifactInventory, ContainerKind, InventoryArtifact, InventoryArtifactKind};
pub use list_times::{list_times_bytes, list_times_path, TimeEntry, TimeListing};
pub use metadata::{
    inspect_metadata_bytes, inspect_metadata_path, MetadataInspection, MetadataProperty,
    MetadataSection,
};
pub use probe::{probe_format_bytes, probe_format_path, FormatProbe};
pub use report_indicators::{DocumentIndicator, IndicatorReport};
/// Structured summary models for CLI and report outputs.
pub use summary::{
    summarize_document, DocumentSummary, MetadataSummary, NodeCount, ParseMetricsSummary,
    SecuritySummary, TextStatsSummary, ThreatIndicatorSummary,
};
pub use vba::{VbaModuleReport, VbaProjectReport, VbaRecognitionReport, VbaRecognitionStatus};

use use_cases::{
    AnalyzeSecurity, AnalyzeSecurityUseCase, DiffDocuments, ParseDocument, ParseDocumentUseCase,
    RunRules, SummarizeUseCase,
};

/// Result alias for all docir-app operations.
pub type AppResult<T> = Result<T, AppError>;

/// High-level error type for application workflows.
#[derive(Debug, Error)]
pub enum AppError {
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

/// Security analysis port for application workflows.
///
/// Implementations compute security conclusions from an existing IR.
pub trait SecurityAnalyzerPort {
    /// Analyze security signals for a parsed document root.
    fn analyze(&mut self, store: &IrStore, root_id: NodeId) -> AnalysisResult;
}

impl SecurityAnalyzerPort for SecurityAnalyzer {
    fn analyze(&mut self, store: &IrStore, root_id: NodeId) -> AnalysisResult {
        self.analyze(store, root_id)
    }
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
    fn format_summary(&self, summary: &DocumentSummary, source: Option<&str>) -> String;
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
        self.parser.parse_file_with_bytes(path)
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
        self.parser.parse_reader_with_bytes(reader)
    }

    /// Serializes a parsed document to JSON.
    pub fn serialize_json(&self, parsed: &ParsedDocument, pretty: bool) -> AppResult<String> {
        self.serializer.to_json(parsed, pretty)
    }

    /// Builds a structured summary for a parsed document.
    pub fn build_summary(&self, parsed: &ParsedDocument) -> Option<DocumentSummary> {
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
    pub fn format_summary(&self, parsed: &ParsedDocument, source: Option<&str>) -> Option<String> {
        self.build_summary(parsed)
            .map(|summary| self.summary_presenter.format_summary(&summary, source))
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
mod tests {
    use super::*;
    use docir_core::ir::{Document, IRNode, Paragraph, Run};
    use docir_core::security::ThreatLevel;
    use std::cell::{Cell, RefCell};
    use std::io::{Cursor, Read, Seek};
    use std::path::Path;
    use std::rc::Rc;

    fn make_parsed_document(format: DocumentFormat) -> ParsedDocument {
        let mut store = IrStore::new();
        let mut doc = Document::new(format);
        let mut paragraph = Paragraph::new();
        let run = Run::new("hello");
        let run_id = run.id;
        paragraph.runs.push(run_id);
        let paragraph_id = paragraph.id;
        doc.content.push(paragraph_id);
        let root_id = doc.id;
        store.insert(IRNode::Run(run));
        store.insert(IRNode::Paragraph(paragraph));
        store.insert(IRNode::Document(doc));

        ParsedDocument::new(ParserParsedDocument {
            root_id,
            format,
            store,
            metrics: Some(ParseMetrics::default()),
        })
    }

    struct MockParser {
        format: DocumentFormat,
        scan_calls: Rc<Cell<u32>>,
        parse_calls: Rc<Cell<u32>>,
    }

    impl ParserPort for MockParser {
        fn parse_file<P: AsRef<Path>>(&self, _path: P) -> AppResult<ParsedDocument> {
            self.parse_calls
                .set(self.parse_calls.get().saturating_add(1));
            Ok(make_parsed_document(self.format))
        }

        fn parse_bytes(&self, _data: &[u8]) -> AppResult<ParsedDocument> {
            self.parse_calls
                .set(self.parse_calls.get().saturating_add(1));
            Ok(make_parsed_document(self.format))
        }

        fn parse_reader<R: Read + Seek>(&self, _reader: R) -> AppResult<ParsedDocument> {
            self.parse_calls
                .set(self.parse_calls.get().saturating_add(1));
            Ok(make_parsed_document(self.format))
        }

        fn parse_file_with_bytes<P: AsRef<Path>>(
            &self,
            _path: P,
        ) -> AppResult<(ParsedDocument, Vec<u8>)> {
            self.parse_calls
                .set(self.parse_calls.get().saturating_add(1));
            Ok((make_parsed_document(self.format), b"file".to_vec()))
        }

        fn parse_reader_with_bytes<R: Read + Seek>(
            &self,
            _reader: R,
        ) -> AppResult<(ParsedDocument, Vec<u8>)> {
            self.parse_calls
                .set(self.parse_calls.get().saturating_add(1));
            Ok((make_parsed_document(self.format), b"reader".to_vec()))
        }
    }

    impl SecurityScannerPort for MockParser {
        fn scan_security_bytes(&self, _data: &[u8], _store: &mut IrStore) -> AppResult<()> {
            self.scan_calls.set(self.scan_calls.get().saturating_add(1));
            Ok(())
        }
    }

    struct NoopEnricher;

    impl SecurityEnricherPort for NoopEnricher {
        fn enrich(&self, _store: &mut IrStore, _root_id: NodeId) {}
    }

    struct JsonSerializer {
        pretty_seen: Rc<RefCell<Vec<bool>>>,
    }

    impl SerializerPort for JsonSerializer {
        fn to_json(&self, parsed: &ParsedDocument, pretty: bool) -> AppResult<String> {
            self.pretty_seen.borrow_mut().push(pretty);
            Ok(format!(
                "{{\"format\":\"{:?}\",\"root\":\"{}\"}}",
                parsed.format(),
                parsed.root_id()
            ))
        }
    }

    struct ConstantAnalyzer {
        calls: Rc<Cell<u32>>,
    }

    impl SecurityAnalyzerPort for ConstantAnalyzer {
        fn analyze(&mut self, _store: &IrStore, _root_id: NodeId) -> AnalysisResult {
            self.calls.set(self.calls.get().saturating_add(1));
            AnalysisResult {
                threat_level: ThreatLevel::Low,
                findings: Vec::new(),
                has_macros: false,
                has_ole_objects: false,
                has_external_refs: false,
                has_dde: false,
                has_xlm_macros: false,
            }
        }
    }

    struct EmptyRulesEngine;

    impl RulesEnginePort for EmptyRulesEngine {
        fn run_with_profile(
            &self,
            _store: &IrStore,
            _root_id: NodeId,
            _profile: &RuleProfile,
        ) -> RuleReport {
            RuleReport {
                findings: Vec::new(),
            }
        }
    }

    #[test]
    fn parsed_document_accessors_return_inner_values() {
        let parsed = make_parsed_document(DocumentFormat::WordProcessing);
        assert_eq!(parsed.format(), DocumentFormat::WordProcessing);
        assert!(parsed.document().is_some());
        assert!(parsed.security_info().is_some());
        assert!(parsed.metrics().is_some());
    }

    #[test]
    fn docir_app_facade_routes_parse_serialize_rules_and_security() {
        let scan_calls = Rc::new(Cell::new(0));
        let parse_calls = Rc::new(Cell::new(0));
        let analyzer_calls = Rc::new(Cell::new(0));
        let pretty_seen = Rc::new(RefCell::new(Vec::new()));

        let parser = MockParser {
            format: DocumentFormat::WordProcessing,
            scan_calls: scan_calls.clone(),
            parse_calls: parse_calls.clone(),
        };

        let app = DocirApp::with_parser_and_ports_and_enricher(
            parser,
            {
                let analyzer_calls = analyzer_calls.clone();
                move || {
                    Box::new(ConstantAnalyzer {
                        calls: analyzer_calls.clone(),
                    })
                }
            },
            Box::new(NoopEnricher),
            || Box::new(EmptyRulesEngine),
            Box::new(JsonSerializer {
                pretty_seen: pretty_seen.clone(),
            }),
            adapters::default_summary_presenter(),
        );

        let parsed_from_file = app
            .parse_file("ignored")
            .expect("parse_file should succeed");
        let _parsed_from_bytes = app
            .parse_bytes(b"bytes")
            .expect("parse_bytes should succeed");
        let _parsed_from_reader = app
            .parse_reader(Cursor::new(b"reader".to_vec()))
            .expect("parse_reader should succeed");

        let json = app
            .serialize_json(&parsed_from_file, true)
            .expect("serialize_json should succeed");
        let report = app.run_rules(&parsed_from_file, &RuleProfile::default());
        let security = app.analyze_security(&parsed_from_file);

        assert!(json.contains("\"format\""));
        assert!(report.is_empty());
        assert_eq!(security.threat_level, ThreatLevel::Low);
        assert_eq!(scan_calls.get(), 3);
        assert_eq!(parse_calls.get(), 3);
        assert_eq!(analyzer_calls.get(), 1);
        assert_eq!(*pretty_seen.borrow(), vec![true]);
    }

    #[test]
    fn diff_reports_no_changes_for_equal_documents() {
        let left = make_parsed_document(DocumentFormat::WordProcessing);
        let right = make_parsed_document(DocumentFormat::WordProcessing);
        let parser = MockParser {
            format: DocumentFormat::WordProcessing,
            scan_calls: Rc::new(Cell::new(0)),
            parse_calls: Rc::new(Cell::new(0)),
        };
        let app = DocirApp::with_parser(parser);
        let diff = app.diff(&left, &right).expect("diff should succeed");
        assert!(diff.added.is_empty());
        assert!(diff.removed.is_empty());
        assert!(diff.modified.is_empty());
    }
}
