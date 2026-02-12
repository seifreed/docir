use crate::{AppResult, ParsedDocument, ParserPort, RulesEnginePort, SecurityAnalyzerPort};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;
use docir_diff::{DiffEngine, DiffResult};
use docir_rules::{RuleProfile, RuleReport};
use docir_security::analyzer::AnalysisResult;
use docir_security::{populate_security_indicators, SecurityAnalyzer};
use docir_serialization::json::to_json;
use std::io::{Read, Seek};
use std::path::Path;

pub(crate) struct ParseDocument<'a, P: ParserPort> {
    parser: &'a P,
}

impl<'a, P: ParserPort> ParseDocument<'a, P> {
    pub(crate) fn new(parser: &'a P) -> Self {
        Self { parser }
    }

    pub(crate) fn parse_file<Pth: AsRef<Path>>(&self, path: Pth) -> AppResult<ParsedDocument> {
        let mut parsed = self.parser.parse_file(path)?;
        EnrichSecurity::run(&mut parsed);
        Ok(parsed)
    }

    pub(crate) fn parse_bytes(&self, data: &[u8]) -> AppResult<ParsedDocument> {
        let mut parsed = self.parser.parse_bytes(data)?;
        EnrichSecurity::run(&mut parsed);
        Ok(parsed)
    }

    pub(crate) fn parse_reader<R: Read + Seek>(&self, reader: R) -> AppResult<ParsedDocument> {
        let mut parsed = self.parser.parse_reader(reader)?;
        EnrichSecurity::run(&mut parsed);
        Ok(parsed)
    }
}

pub(crate) struct EnrichSecurity;

impl EnrichSecurity {
    pub(crate) fn run(parsed: &mut ParsedDocument) {
        let root_id = parsed.root_id();
        populate_security_indicators(parsed.store_mut(), root_id);
    }
}

pub(crate) struct SerializeDocument;

impl SerializeDocument {
    pub(crate) fn to_json(parsed: &ParsedDocument, pretty: bool) -> AppResult<String> {
        Ok(to_json(parsed.store(), parsed.root_id(), pretty)?)
    }
}

pub(crate) struct AnalyzeSecurity<F>
where
    F: Fn() -> Box<dyn SecurityAnalyzerPort>,
{
    analyzer_factory: F,
}

impl<F> AnalyzeSecurity<F>
where
    F: Fn() -> Box<dyn SecurityAnalyzerPort>,
{
    pub(crate) fn new(analyzer_factory: F) -> Self {
        Self { analyzer_factory }
    }

    pub(crate) fn run(&self, store: &IrStore, root_id: NodeId) -> AnalysisResult {
        let mut analyzer = (self.analyzer_factory)();
        analyzer.analyze(store, root_id)
    }
}

pub(crate) struct RunRules<F>
where
    F: Fn() -> Box<dyn RulesEnginePort>,
{
    engine_factory: F,
}

impl<F> RunRules<F>
where
    F: Fn() -> Box<dyn RulesEnginePort>,
{
    pub(crate) fn new(engine_factory: F) -> Self {
        Self { engine_factory }
    }

    pub(crate) fn run(
        &self,
        store: &IrStore,
        root_id: NodeId,
        profile: &RuleProfile,
    ) -> RuleReport {
        let engine = (self.engine_factory)();
        engine.run_with_profile(store, root_id, profile)
    }
}

pub(crate) struct DiffDocuments;

impl DiffDocuments {
    pub(crate) fn diff(left: &ParsedDocument, right: &ParsedDocument) -> DiffResult {
        DiffEngine::diff(left.store(), left.root_id(), right.store(), right.root_id())
    }
}

pub(crate) struct DefaultSecurityAnalyzerFactory;

impl DefaultSecurityAnalyzerFactory {
    pub(crate) fn build() -> Box<dyn SecurityAnalyzerPort> {
        Box::new(SecurityAnalyzer::new())
    }
}
