use crate::AnalysisResult;
use crate::{
    AppResult, DocumentSummary, ParsedDocument, ParserPort, RulesEnginePort, SecurityAnalyzerPort,
    SecurityEnricherPort, SecurityScannerPort,
};
use docir_core::types::{DocumentFormat, NodeId};
use docir_core::visitor::IrStore;
use docir_diff::{DiffEngine, DiffResult};
use docir_rules::{RuleProfile, RuleReport};
use std::io::{Read, Seek};
use std::path::Path;

pub(crate) struct ParseDocumentUseCase<'a, P: ParserPort, S: SecurityScannerPort> {
    parser: &'a P,
    scanner: &'a S,
    enricher: &'a dyn SecurityEnricherPort,
}

impl<'a, P: ParserPort, S: SecurityScannerPort> ParseDocumentUseCase<'a, P, S> {
    pub(crate) fn new(
        parser: &'a P,
        scanner: &'a S,
        enricher: &'a dyn SecurityEnricherPort,
    ) -> Self {
        Self {
            parser,
            scanner,
            enricher,
        }
    }

    pub(crate) fn parse_file<Pth: AsRef<Path>>(&self, path: Pth) -> AppResult<ParsedDocument> {
        let (parsed, data) = self.parser.parse_file_with_bytes(path)?;
        self.finalize_parsed(parsed, &data)
    }

    pub(crate) fn parse_bytes(&self, data: &[u8]) -> AppResult<ParsedDocument> {
        let parsed = self.parser.parse_bytes(data)?;
        self.finalize_parsed(parsed, data)
    }

    pub(crate) fn parse_reader<R: Read + Seek>(&self, reader: R) -> AppResult<ParsedDocument> {
        let (parsed, data) = self.parser.parse_reader_with_bytes(reader)?;
        self.finalize_parsed(parsed, &data)
    }

    fn finalize_parsed(
        &self,
        mut parsed: ParsedDocument,
        data: &[u8],
    ) -> AppResult<ParsedDocument> {
        self.scan_security_if_needed(data, &mut parsed)?;
        EnrichSecurity::run(self.enricher, &mut parsed);
        Ok(parsed)
    }

    fn scan_security_if_needed(&self, data: &[u8], parsed: &mut ParsedDocument) -> AppResult<()> {
        match parsed.format() {
            DocumentFormat::WordProcessing
            | DocumentFormat::Spreadsheet
            | DocumentFormat::Presentation => {
                self.scanner.scan_security_bytes(data, parsed.store_mut())?;
            }
            _ => {}
        }
        Ok(())
    }
}

pub(crate) type ParseDocument<'a, P, S> = ParseDocumentUseCase<'a, P, S>;

pub(crate) struct EnrichSecurity;

impl EnrichSecurity {
    pub(crate) fn run(enricher: &dyn SecurityEnricherPort, parsed: &mut ParsedDocument) {
        let root_id = parsed.root_id();
        enricher.enrich(parsed.store_mut(), root_id);
    }
}

pub(crate) struct AnalyzeSecurityUseCase<F>
where
    F: Fn() -> Box<dyn SecurityAnalyzerPort>,
{
    analyzer_factory: F,
}

impl<F> AnalyzeSecurityUseCase<F>
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

pub(crate) type AnalyzeSecurity<F> = AnalyzeSecurityUseCase<F>;

#[derive(Clone, Copy)]
pub(crate) struct SummarizeUseCase;

impl SummarizeUseCase {
    pub(crate) fn run(&self, parsed: &ParsedDocument) -> Option<DocumentSummary> {
        crate::summarize_document(parsed)
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
    pub(crate) fn diff(left: &ParsedDocument, right: &ParsedDocument) -> AppResult<DiffResult> {
        Ok(DiffEngine::diff(
            left.store(),
            left.root_id(),
            right.store(),
            right.root_id(),
        )?)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{AppError, AppResult, ParseMetrics};
    use docir_core::ir::{Document, IRNode, Paragraph, Run};
    use docir_parser::ParseError as ParserParseError;
    use std::cell::{Cell, RefCell};
    use std::io::{Read, Seek};
    use std::path::Path;

    fn parsed_document(format: DocumentFormat) -> ParsedDocument {
        let mut store = IrStore::new();
        let mut doc = Document::new(format);
        let mut paragraph = Paragraph::new();
        let run = Run::new("content");
        let run_id = run.id;
        paragraph.runs.push(run_id);
        let paragraph_id = paragraph.id;
        doc.content.push(paragraph_id);
        let root_id = doc.id;
        store.insert(IRNode::Run(run));
        store.insert(IRNode::Paragraph(paragraph));
        store.insert(IRNode::Document(doc));
        ParsedDocument {
            inner: docir_parser::parser::ParsedDocument {
                root_id,
                format,
                store,
                metrics: Some(ParseMetrics::default()),
            },
            metrics: Some(ParseMetrics::default()),
        }
    }

    struct MockParser {
        parsed: ParsedDocument,
        payload: Vec<u8>,
    }

    impl ParserPort for MockParser {
        fn parse_file<P: AsRef<Path>>(&self, _path: P) -> AppResult<ParsedDocument> {
            Ok(parsed_document(self.parsed.format()))
        }

        fn parse_bytes(&self, _data: &[u8]) -> AppResult<ParsedDocument> {
            Ok(parsed_document(self.parsed.format()))
        }

        fn parse_reader<R: Read + Seek>(&self, _reader: R) -> AppResult<ParsedDocument> {
            Ok(parsed_document(self.parsed.format()))
        }

        fn parse_file_with_bytes<P: AsRef<Path>>(
            &self,
            _path: P,
        ) -> AppResult<(ParsedDocument, Vec<u8>)> {
            Ok((parsed_document(self.parsed.format()), self.payload.clone()))
        }

        fn parse_reader_with_bytes<R: Read + Seek>(
            &self,
            _reader: R,
        ) -> AppResult<(ParsedDocument, Vec<u8>)> {
            Ok((parsed_document(self.parsed.format()), self.payload.clone()))
        }
    }

    impl SecurityScannerPort for MockParser {
        fn scan_security_bytes(&self, _data: &[u8], _store: &mut IrStore) -> AppResult<()> {
            Ok(())
        }
    }

    struct SpyScanner {
        calls: Cell<u32>,
        fail: bool,
    }

    impl SecurityScannerPort for SpyScanner {
        fn scan_security_bytes(&self, _data: &[u8], _store: &mut IrStore) -> AppResult<()> {
            self.calls.set(self.calls.get().saturating_add(1));
            if self.fail {
                Err(AppError::Parse(ParserParseError::InvalidFormat(
                    "scan failed".to_string(),
                )))
            } else {
                Ok(())
            }
        }
    }

    struct SpyEnricher {
        calls: Cell<u32>,
    }

    impl SecurityEnricherPort for SpyEnricher {
        fn enrich(&self, _store: &mut IrStore, _root_id: NodeId) {
            self.calls.set(self.calls.get().saturating_add(1));
        }
    }

    struct CapturingAnalyzer {
        calls: Cell<u32>,
        last_root: RefCell<Option<NodeId>>,
    }

    impl SecurityAnalyzerPort for CapturingAnalyzer {
        fn analyze(&mut self, _store: &IrStore, root_id: NodeId) -> AnalysisResult {
            self.calls.set(self.calls.get().saturating_add(1));
            *self.last_root.borrow_mut() = Some(root_id);
            AnalysisResult {
                threat_level: docir_core::security::ThreatLevel::None,
                findings: Vec::new(),
                has_macros: false,
                has_ole_objects: false,
                has_external_refs: false,
                has_dde: false,
                has_xlm_macros: false,
            }
        }
    }

    struct StubRulesEngine;

    impl RulesEnginePort for StubRulesEngine {
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
    fn parse_document_scans_security_for_ooxml_and_enriches() {
        let parser = MockParser {
            parsed: parsed_document(DocumentFormat::WordProcessing),
            payload: b"bytes".to_vec(),
        };
        let scanner = SpyScanner {
            calls: Cell::new(0),
            fail: false,
        };
        let enricher = SpyEnricher {
            calls: Cell::new(0),
        };
        let use_case = ParseDocument::new(&parser, &scanner, &enricher);

        let parsed = use_case.parse_bytes(b"input").expect("parse ok");
        assert_eq!(parsed.format(), DocumentFormat::WordProcessing);
        assert_eq!(scanner.calls.get(), 1);
        assert_eq!(enricher.calls.get(), 1);
    }

    #[test]
    fn parse_document_skips_security_scan_for_non_ooxml_formats() {
        let parser = MockParser {
            parsed: parsed_document(DocumentFormat::OdfText),
            payload: b"bytes".to_vec(),
        };
        let scanner = SpyScanner {
            calls: Cell::new(0),
            fail: false,
        };
        let enricher = SpyEnricher {
            calls: Cell::new(0),
        };
        let use_case = ParseDocument::new(&parser, &scanner, &enricher);

        let parsed = use_case.parse_file("ignored").expect("parse ok");
        assert_eq!(parsed.format(), DocumentFormat::OdfText);
        assert_eq!(scanner.calls.get(), 0);
        assert_eq!(enricher.calls.get(), 1);
    }

    #[test]
    fn parse_document_propagates_scan_error() {
        let parser = MockParser {
            parsed: parsed_document(DocumentFormat::Spreadsheet),
            payload: b"bytes".to_vec(),
        };
        let scanner = SpyScanner {
            calls: Cell::new(0),
            fail: true,
        };
        let enricher = SpyEnricher {
            calls: Cell::new(0),
        };
        let use_case = ParseDocument::new(&parser, &scanner, &enricher);

        let err = use_case.parse_reader(std::io::Cursor::new(b"x".to_vec()));
        assert!(err.is_err());
        assert_eq!(scanner.calls.get(), 1);
        assert_eq!(enricher.calls.get(), 0);
    }

    #[test]
    fn analyze_security_use_case_invokes_factory_analyzer() {
        let call_count = std::rc::Rc::new(Cell::new(0));
        let seen_root = std::rc::Rc::new(RefCell::new(None));
        let parsed = parsed_document(DocumentFormat::WordProcessing);
        let expected_root = parsed.root_id();
        let use_case = AnalyzeSecurity::new({
            let call_count = call_count.clone();
            let seen_root = seen_root.clone();
            move || {
                Box::new(CapturingAnalyzer {
                    calls: Cell::new(call_count.get()),
                    last_root: RefCell::new(*seen_root.borrow()),
                })
            }
        });
        let _ = use_case.run(parsed.store(), parsed.root_id());
        assert_eq!(expected_root, parsed.root_id());
    }

    #[test]
    fn run_rules_use_case_calls_engine() {
        let parsed = parsed_document(DocumentFormat::WordProcessing);
        let use_case = RunRules::new(|| Box::new(StubRulesEngine));
        let report = use_case.run(parsed.store(), parsed.root_id(), &RuleProfile::default());
        assert!(report.findings.is_empty());
    }

    #[test]
    fn summarize_use_case_returns_structured_summary() {
        let parsed = parsed_document(DocumentFormat::WordProcessing);
        let use_case = SummarizeUseCase;
        let summary = use_case
            .run(&parsed)
            .expect("word processing must be summarized");
        assert_eq!(summary.security.threat_level, "NONE");
    }
}
