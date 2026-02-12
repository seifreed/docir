//! Application-level workflows for docir.

use anyhow::Result;
use docir_diff::{DiffEngine, DiffResult};
use docir_parser::parser::ParsedDocument;
use docir_parser::DocumentParser;
pub use docir_parser::ParserConfig;
use docir_rules::{RuleEngine, RuleProfile, RuleReport};
use docir_security::analyzer::AnalysisResult;
use docir_security::SecurityAnalyzer;
use docir_serialization::json::to_json;
use std::io::{Read, Seek};
use std::path::Path;

/// Application facade for docir workflows.
pub struct DocirApp {
    parser: DocumentParser,
}

impl DocirApp {
    /// Creates a new app instance with the provided parser config.
    pub fn new(config: ParserConfig) -> Self {
        Self {
            parser: DocumentParser::with_config(config),
        }
    }

    /// Parses a file from disk.
    pub fn parse_file<P: AsRef<Path>>(&self, path: P) -> Result<ParsedDocument> {
        Ok(self.parser.parse_file(path)?)
    }

    /// Parses from bytes.
    pub fn parse_bytes(&self, data: &[u8]) -> Result<ParsedDocument> {
        Ok(self.parser.parse_bytes(data)?)
    }

    /// Parses from a reader.
    pub fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument> {
        Ok(self.parser.parse_reader(reader)?)
    }

    /// Serializes a parsed document to JSON.
    pub fn serialize_json(&self, parsed: &ParsedDocument, pretty: bool) -> Result<String> {
        Ok(to_json(&parsed.store, parsed.root_id, pretty)?)
    }

    /// Runs security analysis for a parsed document.
    pub fn analyze_security(&self, parsed: &ParsedDocument) -> AnalysisResult {
        let mut analyzer = SecurityAnalyzer::new();
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
