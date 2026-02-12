//! Application-level workflows for docir.

use docir_core::ir::Document;
use docir_core::security::SecurityInfo;
use docir_core::types::{DocumentFormat, NodeId};
use docir_core::visitor::IrStore;
use docir_diff::DiffResult;
use docir_parser::parser::ParseMetrics as ParserParseMetrics;
use docir_parser::parser::ParsedDocument as ParserParsedDocument;
pub use docir_parser::ParserConfig;
use docir_parser::{scan_security_bytes, DocumentParser, ParseError};
pub use docir_rules::RuleProfile;
use docir_rules::{RuleEngine, RuleReport};
use docir_security::analyzer::AnalysisResult;
use docir_security::populate_security_indicators;
use docir_security::SecurityAnalyzer;
use docir_serialization::SerializationError;
use std::io::{Read, Seek};
use std::path::Path;
use thiserror::Error;

mod summary;
mod use_cases;

pub use summary::{
    summarize_document, DocumentSummary, MetadataSummary, NodeCount, ParseMetricsSummary,
    SecuritySummary, TextStatsSummary, ThreatIndicatorSummary,
};

use use_cases::{
    AnalyzeSecurity, DefaultSecurityAnalyzerFactory, DiffDocuments, ParseDocument, RunRules,
    SerializeDocument,
};

pub type AppResult<T> = Result<T, AppError>;

#[derive(Debug, Error)]
pub enum AppError {
    #[error(transparent)]
    Parse(#[from] ParseError),
    #[error(transparent)]
    Serialization(#[from] SerializationError),
}

/// Application-level parsed document wrapper.
#[derive(Debug)]
pub struct ParsedDocument {
    inner: ParserParsedDocument,
    metrics: Option<ParseMetrics>,
}

/// Application-level parse metrics.
pub type ParseMetrics = ParserParseMetrics;

impl ParsedDocument {
    pub(crate) fn new(inner: ParserParsedDocument) -> Self {
        let metrics = inner.metrics.clone();
        Self { inner, metrics }
    }

    pub fn root_id(&self) -> NodeId {
        self.inner.root_id
    }

    pub fn format(&self) -> DocumentFormat {
        self.inner.format
    }

    pub fn store(&self) -> &IrStore {
        &self.inner.store
    }

    pub fn store_mut(&mut self) -> &mut IrStore {
        &mut self.inner.store
    }

    pub fn document(&self) -> Option<&Document> {
        self.inner.document()
    }

    pub fn security_info(&self) -> Option<&SecurityInfo> {
        self.inner.security_info()
    }

    pub fn metrics(&self) -> Option<&ParseMetrics> {
        self.metrics.as_ref()
    }

    pub(crate) fn into_inner(self) -> ParserParsedDocument {
        self.inner
    }
}
/// Parser port for application workflows.
pub trait ParserPort {
    fn parse_file<P: AsRef<Path>>(&self, path: P) -> AppResult<ParsedDocument>;
    fn parse_bytes(&self, data: &[u8]) -> AppResult<ParsedDocument>;
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> AppResult<ParsedDocument>;
    fn parse_file_with_bytes<P: AsRef<Path>>(
        &self,
        path: P,
    ) -> AppResult<(ParsedDocument, Vec<u8>)>;
    fn parse_reader_with_bytes<R: Read + Seek>(
        &self,
        reader: R,
    ) -> AppResult<(ParsedDocument, Vec<u8>)>;
}

impl ParserPort for DocumentParser {
    fn parse_file<P: AsRef<Path>>(&self, path: P) -> AppResult<ParsedDocument> {
        self.parse_file(path)
            .map(ParsedDocument::new)
            .map_err(Into::into)
    }

    fn parse_bytes(&self, data: &[u8]) -> AppResult<ParsedDocument> {
        self.parse_bytes(data)
            .map(ParsedDocument::new)
            .map_err(Into::into)
    }

    fn parse_reader<R: Read + Seek>(&self, reader: R) -> AppResult<ParsedDocument> {
        self.parse_reader(reader)
            .map(ParsedDocument::new)
            .map_err(Into::into)
    }

    fn parse_file_with_bytes<P: AsRef<Path>>(
        &self,
        path: P,
    ) -> AppResult<(ParsedDocument, Vec<u8>)> {
        self.parse_file_with_bytes(path)
            .map(|(parsed, data)| (ParsedDocument::new(parsed), data))
            .map_err(Into::into)
    }

    fn parse_reader_with_bytes<R: Read + Seek>(
        &self,
        reader: R,
    ) -> AppResult<(ParsedDocument, Vec<u8>)> {
        self.parse_reader_with_bytes(reader)
            .map(|(parsed, data)| (ParsedDocument::new(parsed), data))
            .map_err(Into::into)
    }
}

/// Security scanning port for application workflows.
pub trait SecurityScannerPort {
    fn scan_security_bytes(&self, data: &[u8], store: &mut IrStore) -> AppResult<()>;
}

pub struct AppParser {
    parser: DocumentParser,
    config: ParserConfig,
}

impl AppParser {
    pub fn new(parser: DocumentParser, config: ParserConfig) -> Self {
        Self { parser, config }
    }
}

impl ParserPort for AppParser {
    fn parse_file<P: AsRef<Path>>(&self, path: P) -> AppResult<ParsedDocument> {
        self.parser
            .parse_file(path)
            .map(ParsedDocument::new)
            .map_err(Into::into)
    }

    fn parse_bytes(&self, data: &[u8]) -> AppResult<ParsedDocument> {
        self.parser
            .parse_bytes(data)
            .map(ParsedDocument::new)
            .map_err(Into::into)
    }

    fn parse_reader<R: Read + Seek>(&self, reader: R) -> AppResult<ParsedDocument> {
        self.parser
            .parse_reader(reader)
            .map(ParsedDocument::new)
            .map_err(Into::into)
    }

    fn parse_file_with_bytes<P: AsRef<Path>>(
        &self,
        path: P,
    ) -> AppResult<(ParsedDocument, Vec<u8>)> {
        self.parser
            .parse_file_with_bytes(path)
            .map(|(parsed, data)| (ParsedDocument::new(parsed), data))
            .map_err(Into::into)
    }

    fn parse_reader_with_bytes<R: Read + Seek>(
        &self,
        reader: R,
    ) -> AppResult<(ParsedDocument, Vec<u8>)> {
        self.parser
            .parse_reader_with_bytes(reader)
            .map(|(parsed, data)| (ParsedDocument::new(parsed), data))
            .map_err(Into::into)
    }
}

impl SecurityScannerPort for AppParser {
    fn scan_security_bytes(&self, data: &[u8], store: &mut IrStore) -> AppResult<()> {
        scan_security_bytes(&self.config, data, store).map_err(Into::into)
    }
}

/// Security analysis port for application workflows.
pub trait SecurityAnalyzerPort {
    fn analyze(&mut self, store: &IrStore, root_id: NodeId) -> AnalysisResult;
}

impl SecurityAnalyzerPort for SecurityAnalyzer {
    fn analyze(&mut self, store: &IrStore, root_id: NodeId) -> AnalysisResult {
        self.analyze(store, root_id)
    }
}

/// Security enrichment port for application workflows.
pub trait SecurityEnricherPort {
    fn enrich(&self, store: &mut IrStore, root_id: NodeId);
}

struct DefaultSecurityEnricher;

impl SecurityEnricherPort for DefaultSecurityEnricher {
    fn enrich(&self, store: &mut IrStore, root_id: NodeId) {
        populate_security_indicators(store, root_id);
    }
}

/// Rules engine port for application workflows.
pub trait RulesEnginePort {
    fn run_with_profile(
        &self,
        store: &IrStore,
        root_id: NodeId,
        profile: &RuleProfile,
    ) -> RuleReport;
}

struct DefaultRulesEngine;

impl RulesEnginePort for DefaultRulesEngine {
    fn run_with_profile(
        &self,
        store: &IrStore,
        root_id: NodeId,
        profile: &RuleProfile,
    ) -> RuleReport {
        let engine = RuleEngine::with_default_rules();
        engine.run_with_profile(store, root_id, profile)
    }
}

/// Serialization port for application workflows.
pub trait SerializerPort {
    fn to_json(&self, parsed: &ParsedDocument, pretty: bool) -> AppResult<String>;
}

struct DefaultJsonSerializer;

impl SerializerPort for DefaultJsonSerializer {
    fn to_json(&self, parsed: &ParsedDocument, pretty: bool) -> AppResult<String> {
        SerializeDocument::to_json(parsed, pretty)
    }
}

/// Application facade for docir workflows.
pub struct DocirApp<P: ParserPort + SecurityScannerPort = AppParser> {
    parser: P,
    security_analyzer_factory: Box<dyn Fn() -> Box<dyn SecurityAnalyzerPort>>,
    security_enricher: Box<dyn SecurityEnricherPort>,
    rules_engine_factory: Box<dyn Fn() -> Box<dyn RulesEnginePort>>,
    serializer: Box<dyn SerializerPort>,
}

impl DocirApp<AppParser> {
    /// Creates a new app instance with the provided parser config.
    pub fn new(config: ParserConfig) -> Self {
        let mut config = config;
        config.scan_security_on_parse = false;
        let parser = DocumentParser::with_config(config.clone());
        Self::with_parser(AppParser::new(parser, config))
    }
}

impl<P: ParserPort + SecurityScannerPort> DocirApp<P> {
    /// Creates a new app instance with a custom parser implementation.
    pub fn with_parser(parser: P) -> Self {
        Self::with_parser_and_ports(
            parser,
            DefaultSecurityAnalyzerFactory::build,
            || Box::new(DefaultRulesEngine),
            Box::new(DefaultJsonSerializer),
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
            || Box::new(DefaultRulesEngine),
            Box::new(DefaultJsonSerializer),
        )
    }

    /// Creates a new app instance with custom ports.
    pub fn with_parser_and_ports<F, R>(
        parser: P,
        security_analyzer_factory: F,
        rules_engine_factory: R,
        serializer: Box<dyn SerializerPort>,
    ) -> Self
    where
        F: Fn() -> Box<dyn SecurityAnalyzerPort> + 'static,
        R: Fn() -> Box<dyn RulesEnginePort> + 'static,
    {
        Self::with_parser_and_ports_and_enricher(
            parser,
            security_analyzer_factory,
            Box::new(DefaultSecurityEnricher),
            rules_engine_factory,
            serializer,
        )
    }

    /// Creates a new app instance with custom ports and security enricher.
    pub fn with_parser_and_ports_and_enricher<F, R>(
        parser: P,
        security_analyzer_factory: F,
        security_enricher: Box<dyn SecurityEnricherPort>,
        rules_engine_factory: R,
        serializer: Box<dyn SerializerPort>,
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
        }
    }

    /// Parses a file from disk.
    pub fn parse_file<Pth: AsRef<Path>>(&self, path: Pth) -> AppResult<ParsedDocument> {
        ParseDocument::new(&self.parser, &self.parser, self.security_enricher.as_ref())
            .parse_file(path)
    }

    /// Parses from bytes.
    pub fn parse_bytes(&self, data: &[u8]) -> AppResult<ParsedDocument> {
        ParseDocument::new(&self.parser, &self.parser, self.security_enricher.as_ref())
            .parse_bytes(data)
    }

    /// Parses from a reader.
    pub fn parse_reader<R: Read + Seek>(&self, reader: R) -> AppResult<ParsedDocument> {
        ParseDocument::new(&self.parser, &self.parser, self.security_enricher.as_ref())
            .parse_reader(reader)
    }

    /// Serializes a parsed document to JSON.
    pub fn serialize_json(&self, parsed: &ParsedDocument, pretty: bool) -> AppResult<String> {
        self.serializer.to_json(parsed, pretty)
    }

    /// Runs security analysis for a parsed document.
    pub fn analyze_security(&self, parsed: &ParsedDocument) -> AnalysisResult {
        AnalyzeSecurity::new(&self.security_analyzer_factory).run(parsed.store(), parsed.root_id())
    }

    /// Runs rules for a parsed document.
    pub fn run_rules(&self, parsed: &ParsedDocument, profile: &RuleProfile) -> RuleReport {
        RunRules::new(&self.rules_engine_factory).run(parsed.store(), parsed.root_id(), profile)
    }

    /// Computes a diff between two parsed documents.
    pub fn diff(&self, left: &ParsedDocument, right: &ParsedDocument) -> DiffResult {
        DiffDocuments::diff(left, right)
    }
}
