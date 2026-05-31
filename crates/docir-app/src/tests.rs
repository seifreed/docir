use super::*;
use docir_core::ir::{Document, IRNode, Paragraph, Run};
use docir_core::security::ThreatLevel;
use std::cell::{Cell, RefCell};
use std::io::{Cursor, Read, Seek};
use std::path::Path;
use std::rc::Rc;

fn hwp_file_header() -> Vec<u8> {
    let mut header = vec![0u8; 40];
    header[..17].copy_from_slice(b"HWP Document File");
    header[32..36].copy_from_slice(&0x0500_0000u32.to_le_bytes());
    header
}

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
fn hwp_default_jscript_parse_failures_are_reported() {
    let header = hwp_file_header();
    let bytes = test_support::build_test_cfb(&[
        ("FileHeader", &header),
        ("Scripts/DefaultJScript", b"\0\0\0\0\x20\0"),
    ]);
    let app = DocirApp::new(ParserConfig::default());
    let parsed = app.parse_bytes(&bytes).expect("hwp parse");

    let reported = parsed.store().iter().any(|(_, node)| {
        matches!(
            node,
            IRNode::Diagnostics(diag)
                if diag.entries.iter().any(|entry| {
                    entry.code == "HWP_SCRIPT_PARSE_FAILED"
                        && entry.path.as_deref() == Some("Scripts/DefaultJScript")
                })
        )
    });
    assert!(reported, "malformed HWP script stream must not be silent");
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
