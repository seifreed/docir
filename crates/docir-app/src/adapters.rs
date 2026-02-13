//! Infrastructure adapters for application ports.

use crate::{
    AppParseError, AppResult, ParsedDocument, ParserConfig, ParserPort, RuleProfile, RuleReport,
    RulesEnginePort, SecurityAnalyzerPort, SecurityEnricherPort, SecurityScannerPort,
    SerializerPort,
};
use docir_core::visitor::IrStore;
use docir_parser::parser::ParsedDocument as ParserParsedDocument;
use docir_parser::{scan_security_bytes as scan_parser_bytes, DocumentParser, ParseError};
use docir_rules::RuleEngine;
use docir_security::populate_security_indicators;
use docir_security::SecurityAnalyzer;
use std::io::{Read, Seek};
use std::path::Path;

use crate::use_cases::SerializeDocument;

/// Parser adapter that bundles a configured parser with its config.
pub struct AppParser {
    parser: DocumentParser,
    config: docir_parser::ParserConfig,
}

struct DefaultSecurityEnricher;

impl SecurityEnricherPort for DefaultSecurityEnricher {
    fn enrich(&self, store: &mut IrStore, root_id: docir_core::types::NodeId) {
        populate_security_indicators(store, root_id);
    }
}

struct DefaultRulesEngine;

impl RulesEnginePort for DefaultRulesEngine {
    fn run_with_profile(
        &self,
        store: &IrStore,
        root_id: docir_core::types::NodeId,
        profile: &RuleProfile,
    ) -> RuleReport {
        let engine = RuleEngine::with_default_rules();
        engine.run_with_profile(store, root_id, profile)
    }
}

struct DefaultJsonSerializer;

impl SerializerPort for DefaultJsonSerializer {
    fn to_json(&self, parsed: &ParsedDocument, pretty: bool) -> AppResult<String> {
        SerializeDocument::to_json(parsed, pretty)
    }
}

pub(crate) fn default_security_enricher() -> Box<dyn SecurityEnricherPort> {
    Box::new(DefaultSecurityEnricher)
}

pub(crate) fn default_security_analyzer_factory() -> impl Fn() -> Box<dyn SecurityAnalyzerPort> {
    || Box::new(SecurityAnalyzer::new())
}

pub(crate) fn default_rules_engine_factory() -> impl Fn() -> Box<dyn RulesEnginePort> {
    || Box::new(DefaultRulesEngine)
}

pub(crate) fn default_json_serializer() -> Box<dyn SerializerPort> {
    Box::new(DefaultJsonSerializer)
}

impl AppParser {
    pub fn new(parser: DocumentParser, config: ParserConfig) -> Self {
        Self { parser, config }
    }

    pub fn with_config(config: ParserConfig) -> Self {
        let parser = DocumentParser::with_config(config.clone());
        Self { parser, config }
    }
}

impl ParserPort for DocumentParser {
    fn parse_file<P: AsRef<Path>>(&self, path: P) -> AppResult<ParsedDocument> {
        wrap_parsed(self.parse_file(path))
    }

    fn parse_bytes(&self, data: &[u8]) -> AppResult<ParsedDocument> {
        wrap_parsed(self.parse_bytes(data))
    }

    fn parse_reader<R: Read + Seek>(&self, reader: R) -> AppResult<ParsedDocument> {
        wrap_parsed(self.parse_reader(reader))
    }

    fn parse_file_with_bytes<P: AsRef<Path>>(
        &self,
        path: P,
    ) -> AppResult<(ParsedDocument, Vec<u8>)> {
        wrap_parsed_with_bytes(self.parse_file_with_bytes(path))
    }

    fn parse_reader_with_bytes<R: Read + Seek>(
        &self,
        reader: R,
    ) -> AppResult<(ParsedDocument, Vec<u8>)> {
        wrap_parsed_with_bytes(self.parse_reader_with_bytes(reader))
    }
}

impl ParserPort for AppParser {
    fn parse_file<P: AsRef<Path>>(&self, path: P) -> AppResult<ParsedDocument> {
        ParserPort::parse_file(&self.parser, path)
    }

    fn parse_bytes(&self, data: &[u8]) -> AppResult<ParsedDocument> {
        ParserPort::parse_bytes(&self.parser, data)
    }

    fn parse_reader<R: Read + Seek>(&self, reader: R) -> AppResult<ParsedDocument> {
        ParserPort::parse_reader(&self.parser, reader)
    }

    fn parse_file_with_bytes<P: AsRef<Path>>(
        &self,
        path: P,
    ) -> AppResult<(ParsedDocument, Vec<u8>)> {
        ParserPort::parse_file_with_bytes(&self.parser, path)
    }

    fn parse_reader_with_bytes<R: Read + Seek>(
        &self,
        reader: R,
    ) -> AppResult<(ParsedDocument, Vec<u8>)> {
        ParserPort::parse_reader_with_bytes(&self.parser, reader)
    }
}

impl SecurityScannerPort for AppParser {
    fn scan_security_bytes(&self, data: &[u8], store: &mut IrStore) -> AppResult<()> {
        scan_security_bytes(&self.config, data, store).map_err(Into::into)
    }
}

fn scan_security_bytes(
    config: &docir_parser::ParserConfig,
    data: &[u8],
    store: &mut IrStore,
) -> Result<(), AppParseError> {
    scan_parser_bytes(config, data, store).map_err(AppParseError::from)
}

fn wrap_parsed(result: Result<ParserParsedDocument, ParseError>) -> AppResult<ParsedDocument> {
    map_parse_error(result.map(ParsedDocument::new))
}

fn wrap_parsed_with_bytes(
    result: Result<(ParserParsedDocument, Vec<u8>), ParseError>,
) -> AppResult<(ParsedDocument, Vec<u8>)> {
    map_parse_error(result.map(|(parsed, data)| (ParsedDocument::new(parsed), data)))
}

fn map_parse_error<T>(result: Result<T, ParseError>) -> AppResult<T> {
    result.map_err(AppParseError::from).map_err(Into::into)
}
