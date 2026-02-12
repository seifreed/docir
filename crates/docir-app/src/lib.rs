//! Application-level workflows for docir.

use anyhow::Result;
use docir_core::ir::Document;
use docir_core::security::SecurityInfo;
use docir_core::types::{DocumentFormat, NodeId};
use docir_core::visitor::IrStore;
use docir_diff::DiffResult;
use docir_parser::parser::ParseMetrics as ParserParseMetrics;
use docir_parser::parser::ParsedDocument as ParserParsedDocument;
pub use docir_parser::ParserConfig;
use docir_parser::{DocumentParser, ParseError};
pub use docir_rules::RuleProfile;
use docir_rules::{RuleEngine, RuleReport};
use docir_security::analyzer::AnalysisResult;
use docir_security::SecurityAnalyzer;
use std::io::{Read, Seek};
use std::path::Path;

mod use_cases;

use use_cases::{
    AnalyzeSecurity, DefaultSecurityAnalyzerFactory, DiffDocuments, ParseDocument, RunRules,
    SerializeDocument,
};

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
    fn parse_file<P: AsRef<Path>>(&self, path: P) -> Result<ParsedDocument, ParseError>;
    fn parse_bytes(&self, data: &[u8]) -> Result<ParsedDocument, ParseError>;
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError>;
}

impl ParserPort for DocumentParser {
    fn parse_file<P: AsRef<Path>>(&self, path: P) -> Result<ParsedDocument, ParseError> {
        self.parse_file(path).map(ParsedDocument::new)
    }

    fn parse_bytes(&self, data: &[u8]) -> Result<ParsedDocument, ParseError> {
        self.parse_bytes(data).map(ParsedDocument::new)
    }

    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        self.parse_reader(reader).map(ParsedDocument::new)
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
    fn to_json(&self, parsed: &ParsedDocument, pretty: bool) -> Result<String>;
}

struct DefaultJsonSerializer;

impl SerializerPort for DefaultJsonSerializer {
    fn to_json(&self, parsed: &ParsedDocument, pretty: bool) -> Result<String> {
        SerializeDocument::to_json(parsed, pretty)
    }
}

/// Application facade for docir workflows.
pub struct DocirApp<P: ParserPort = DocumentParser> {
    parser: P,
    security_analyzer_factory: Box<dyn Fn() -> Box<dyn SecurityAnalyzerPort>>,
    rules_engine_factory: Box<dyn Fn() -> Box<dyn RulesEnginePort>>,
    serializer: Box<dyn SerializerPort>,
}

impl DocirApp<DocumentParser> {
    /// Creates a new app instance with the provided parser config.
    pub fn new(config: ParserConfig) -> Self {
        Self::with_parser(DocumentParser::with_config(config))
    }
}

impl<P: ParserPort> DocirApp<P> {
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
        Self {
            parser,
            security_analyzer_factory: Box::new(security_analyzer_factory),
            rules_engine_factory: Box::new(rules_engine_factory),
            serializer,
        }
    }

    /// Parses a file from disk.
    pub fn parse_file<Pth: AsRef<Path>>(&self, path: Pth) -> Result<ParsedDocument> {
        ParseDocument::new(&self.parser).parse_file(path)
    }

    /// Parses from bytes.
    pub fn parse_bytes(&self, data: &[u8]) -> Result<ParsedDocument> {
        ParseDocument::new(&self.parser).parse_bytes(data)
    }

    /// Parses from a reader.
    pub fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument> {
        ParseDocument::new(&self.parser).parse_reader(reader)
    }

    /// Serializes a parsed document to JSON.
    pub fn serialize_json(&self, parsed: &ParsedDocument, pretty: bool) -> Result<String> {
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
