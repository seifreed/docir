//! Application-level workflows for docir.

use anyhow::Result;
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;
use docir_diff::{DiffEngine, DiffResult};
use docir_parser::parser::ParsedDocument;
pub use docir_parser::ParserConfig;
use docir_parser::{DocumentParser, ParseError};
use docir_rules::{RuleEngine, RuleProfile, RuleReport};
use docir_security::analyzer::AnalysisResult;
use docir_security::{populate_security_indicators, SecurityAnalyzer};
use docir_serialization::json::to_json;
use std::io::{Read, Seek};
use std::path::Path;

/// Parser port for application workflows.
pub trait ParserPort {
    fn parse_file<P: AsRef<Path>>(&self, path: P) -> Result<ParsedDocument, ParseError>;
    fn parse_bytes(&self, data: &[u8]) -> Result<ParsedDocument, ParseError>;
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError>;
}

impl ParserPort for DocumentParser {
    fn parse_file<P: AsRef<Path>>(&self, path: P) -> Result<ParsedDocument, ParseError> {
        self.parse_file(path)
    }

    fn parse_bytes(&self, data: &[u8]) -> Result<ParsedDocument, ParseError> {
        self.parse_bytes(data)
    }

    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        self.parse_reader(reader)
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

/// Application facade for docir workflows.
pub struct DocirApp<P: ParserPort = DocumentParser> {
    parser: P,
    security_analyzer_factory: Box<dyn Fn() -> Box<dyn SecurityAnalyzerPort>>,
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
        Self::with_parser_and_security(parser, || Box::new(SecurityAnalyzer::new()))
    }

    /// Creates a new app instance with custom parser and security analyzer factory.
    pub fn with_parser_and_security<F>(parser: P, security_analyzer_factory: F) -> Self
    where
        F: Fn() -> Box<dyn SecurityAnalyzerPort> + 'static,
    {
        Self {
            parser,
            security_analyzer_factory: Box::new(security_analyzer_factory),
        }
    }

    /// Parses a file from disk.
    pub fn parse_file<Pth: AsRef<Path>>(&self, path: Pth) -> Result<ParsedDocument> {
        let mut parsed = self.parser.parse_file(path)?;
        populate_security_indicators(&mut parsed.store, parsed.root_id);
        Ok(parsed)
    }

    /// Parses from bytes.
    pub fn parse_bytes(&self, data: &[u8]) -> Result<ParsedDocument> {
        let mut parsed = self.parser.parse_bytes(data)?;
        populate_security_indicators(&mut parsed.store, parsed.root_id);
        Ok(parsed)
    }

    /// Parses from a reader.
    pub fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument> {
        let mut parsed = self.parser.parse_reader(reader)?;
        populate_security_indicators(&mut parsed.store, parsed.root_id);
        Ok(parsed)
    }

    /// Serializes a parsed document to JSON.
    pub fn serialize_json(&self, parsed: &ParsedDocument, pretty: bool) -> Result<String> {
        Ok(to_json(&parsed.store, parsed.root_id, pretty)?)
    }

    /// Runs security analysis for a parsed document.
    pub fn analyze_security(&self, parsed: &ParsedDocument) -> AnalysisResult {
        let mut analyzer = (self.security_analyzer_factory)();
        analyzer.analyze(&parsed.store, parsed.root_id)
    }

    /// Runs rules for a parsed document.
    pub fn run_rules(&self, parsed: &ParsedDocument, profile: &RuleProfile) -> RuleReport {
        let engine = RuleEngine::with_default_rules();
        engine.run_with_profile(&parsed.store, parsed.root_id, profile)
    }

    /// Computes a diff between two parsed documents.
    pub fn diff(&self, left: &ParsedDocument, right: &ParsedDocument) -> DiffResult {
        DiffEngine::diff(&left.store, left.root_id, &right.store, right.root_id)
    }
}
