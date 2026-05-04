//! Infrastructure adapters for application ports.

use crate::ports::CfbStreamReaderPort;
use crate::summary::DocumentSummary;
use crate::{
    AppError, AppResult, ParsedDocument, ParserConfig, ParserPort, RuleProfile, RuleReport,
    RulesEnginePort, SecurityAnalyzerPort, SecurityEnricherPort, SecurityScannerPort,
    SerializerPort, SummaryPresenterPort,
};
use docir_core::visitor::IrStore;
use docir_parser::ole::Cfb;
use docir_parser::parser::ParsedDocument as ParserParsedDocument;
use docir_parser::{scan_security_bytes as scan_parser_bytes, DocumentParser, ParseError};
use docir_rules::RuleEngine;
use docir_security::populate_security_indicators;
use docir_security::SecurityAnalyzer;
use docir_serialization::json::to_json;
use std::io::{Read, Seek};
use std::path::Path;

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
        Ok(to_json(parsed.store(), parsed.root_id(), pretty)?)
    }
}

struct DefaultSummaryPresenter;

pub(crate) struct ParserCfbStreamReader;

impl CfbStreamReaderPort for ParserCfbStreamReader {
    fn read_streams(
        &self,
        data: &[u8],
        stream_names: &[&str],
    ) -> AppResult<Vec<(String, Vec<u8>)>> {
        let cfb = Cfb::parse(data.to_vec())?;
        Ok(stream_names
            .iter()
            .filter_map(|name| {
                cfb.read_stream(name)
                    .map(|bytes| ((*name).to_string(), bytes))
            })
            .collect())
    }
}

impl SummaryPresenterPort for DefaultSummaryPresenter {
    fn format_summary(&self, summary: &DocumentSummary, source: Option<&str>) -> String {
        let mut out = String::new();
        let total_nodes: usize = summary.node_counts.iter().map(|entry| entry.count).sum();
        out.push_str("Document Summary\n");
        out.push_str("================\n\n");
        if let Some(source) = source {
            out.push_str(&format!("File: {}\n", source));
        }
        out.push_str(&format!("Format: {}\n", summary.format));
        out.push_str(&format!("Nodes: {}\n\n", total_nodes));
        format_metadata(&mut out, summary);
        format_structure(&mut out, summary);
        format_text_stats(&mut out, summary);
        format_metrics(&mut out, summary);
        format_security(&mut out, summary);
        format_threat_indicators(&mut out, summary);
        out
    }
}

fn format_metadata(output: &mut String, summary: &DocumentSummary) {
    if summary.metadata.title.is_none()
        && summary.metadata.author.is_none()
        && summary.metadata.modified.is_none()
        && summary.metadata.application.is_none()
    {
        return;
    }

    output.push_str("Metadata:\n");
    if let Some(title) = &summary.metadata.title {
        output.push_str(&format!("  Title: {}\n", title));
    }
    if let Some(creator) = &summary.metadata.author {
        output.push_str(&format!("  Author: {}\n", creator));
    }
    if let Some(modified) = &summary.metadata.modified {
        output.push_str(&format!("  Modified: {}\n", modified));
    }
    if let Some(app) = &summary.metadata.application {
        output.push_str(&format!("  Application: {}\n", app));
    }
    output.push('\n');
}

fn format_structure(output: &mut String, summary: &DocumentSummary) {
    output.push_str("Structure:\n");
    for count in &summary.node_counts {
        if count.count > 0 {
            output.push_str(&format!("  {}: {}\n", count.node_type, count.count));
        }
    }
    output.push('\n');
}

fn format_text_stats(output: &mut String, summary: &DocumentSummary) {
    output.push_str("Text Statistics:\n");
    output.push_str(&format!(
        "  Characters: {}\n",
        summary.text_stats.char_count
    ));
    output.push_str(&format!("  Words: ~{}\n", summary.text_stats.word_count));
    output.push('\n');
}

fn format_metrics(output: &mut String, summary: &DocumentSummary) {
    let Some(metrics) = &summary.metrics else {
        return;
    };

    output.push_str("Parse Metrics (ms):\n");
    output.push_str(&format!("  Content Types: {}\n", metrics.content_types_ms));
    output.push_str(&format!("  Relationships: {}\n", metrics.relationships_ms));
    output.push_str(&format!("  Main Parse: {}\n", metrics.main_parse_ms));
    output.push_str(&format!("  Shared Parts: {}\n", metrics.shared_parts_ms));
    output.push_str(&format!("  Security Scan: {}\n", metrics.security_scan_ms));
    output.push_str(&format!(
        "  Extension Parts: {}\n",
        metrics.extension_parts_ms
    ));
    output.push_str(&format!(
        "  Normalization: {}\n\n",
        metrics.normalization_ms
    ));
}

fn format_security(output: &mut String, summary: &DocumentSummary) {
    output.push_str("Security:\n");
    output.push_str(&format!(
        "  Threat Level: {}\n",
        summary.security.threat_level
    ));
    output.push_str(&format!(
        "  VBA Macros: {}\n",
        if summary.security.has_macro_project {
            "YES - DETECTED"
        } else {
            "No"
        }
    ));
    output.push_str(&format!(
        "  OLE Objects: {}\n",
        count_or_none(summary.security.ole_objects)
    ));
    output.push_str(&format!(
        "  External References: {}\n",
        count_or_none(summary.security.external_refs)
    ));
    output.push_str(&format!(
        "  DDE Fields: {}\n",
        count_or_none(summary.security.dde_fields)
    ));
    output.push_str(&format!(
        "  ActiveX Controls: {}\n",
        count_or_none(summary.security.activex_controls)
    ));
    output.push_str(&format!(
        "  XLM Macros: {}\n",
        count_or_none(summary.security.xlm_macros)
    ));
    output.push('\n');
}

fn format_threat_indicators(output: &mut String, summary: &DocumentSummary) {
    if summary.threat_indicators.is_empty() {
        return;
    }

    output.push_str("Threat Indicators:\n");
    for indicator in &summary.threat_indicators {
        output.push_str(&format!(
            "  [{}] {:?}: {}\n",
            indicator.severity, indicator.indicator_type, indicator.description
        ));
    }
}

fn count_or_none(count: usize) -> String {
    if count == 0 {
        "No".to_string()
    } else {
        format!("{} found", count)
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

pub(crate) fn default_summary_presenter() -> Box<dyn SummaryPresenterPort> {
    Box::new(DefaultSummaryPresenter)
}

pub(crate) fn default_cfb_stream_reader() -> ParserCfbStreamReader {
    ParserCfbStreamReader
}

impl AppParser {
    /// Public API entrypoint: new.
    pub fn new(parser: DocumentParser, config: ParserConfig) -> Self {
        Self { parser, config }
    }

    /// Public API entrypoint: with_config.
    pub fn with_config(config: ParserConfig) -> Self {
        let parser = DocumentParser::with_config(config.clone());
        Self { parser, config }
    }

    pub(crate) fn zip_config(&self) -> &docir_parser::zip_handler::ZipConfig {
        &self.config.zip_config
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
        scan_security_bytes(&self.config, data, store)
    }
}

fn scan_security_bytes(
    config: &docir_parser::ParserConfig,
    data: &[u8],
    store: &mut IrStore,
) -> Result<(), AppError> {
    scan_parser_bytes(config, data, store).map_err(AppError::from)
}

fn wrap_parsed(result: Result<ParserParsedDocument, ParseError>) -> AppResult<ParsedDocument> {
    map_parsed_result(result, ParsedDocument::new)
}

fn wrap_parsed_with_bytes(
    result: Result<(ParserParsedDocument, Vec<u8>), ParseError>,
) -> AppResult<(ParsedDocument, Vec<u8>)> {
    map_parsed_result(result, |(parsed, data)| (ParsedDocument::new(parsed), data))
}

fn map_parsed_result<T, U, F>(result: Result<T, ParseError>, map: F) -> AppResult<U>
where
    F: FnOnce(T) -> U,
{
    result.map(map).map_err(AppError::from)
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::{Document, IRNode, Paragraph, Run};
    use docir_core::types::DocumentFormat;
    use std::fs;
    use std::io::Cursor;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn parser_parsed_document(format: DocumentFormat) -> ParserParsedDocument {
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

        ParserParsedDocument {
            root_id,
            format,
            store,
            metrics: None,
        }
    }

    fn repo_root() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .expect("workspace crate dir")
            .parent()
            .expect("workspace root")
            .to_path_buf()
    }

    fn fixture_bytes(path: &str) -> Vec<u8> {
        fs::read(repo_root().join(path)).expect("read fixture bytes")
    }

    #[test]
    fn map_parsed_result_maps_success_value() {
        let mapped = map_parsed_result::<u32, String, _>(Ok(7), |v| format!("v={v}"))
            .expect("mapped result");
        assert_eq!(mapped, "v=7");
    }

    #[test]
    fn map_parsed_result_converts_parse_errors() {
        let err = map_parsed_result::<u32, u32, _>(
            Err(ParseError::UnsupportedFormat("x".to_string())),
            |v| v,
        )
        .expect_err("must convert parse error");
        assert!(format!("{err}").contains("Unsupported format"));
    }

    #[test]
    fn wrap_parsed_and_wrap_parsed_with_bytes_keep_payload() {
        let wrapped = wrap_parsed(Ok(parser_parsed_document(DocumentFormat::WordProcessing)))
            .expect("wrapped parse");
        assert_eq!(wrapped.format(), DocumentFormat::WordProcessing);
        assert!(wrapped.document().is_some());

        let payload = b"abc".to_vec();
        let (wrapped, bytes) = wrap_parsed_with_bytes(Ok((
            parser_parsed_document(DocumentFormat::Spreadsheet),
            payload.clone(),
        )))
        .expect("wrapped parse with bytes");
        assert_eq!(wrapped.format(), DocumentFormat::Spreadsheet);
        assert_eq!(bytes, payload);
    }

    #[test]
    fn default_json_serializer_outputs_document_json() {
        let serializer = default_json_serializer();
        let parsed = ParsedDocument::new(parser_parsed_document(DocumentFormat::WordProcessing));
        let json = serializer.to_json(&parsed, true).expect("serialize json");
        assert!(json.contains("\"type\": \"Document\""));
    }

    #[test]
    fn app_parser_with_config_parses_and_scans_fixture_bytes() {
        let parser = AppParser::with_config(ParserConfig::default());
        let bytes = fixture_bytes("fixtures/ooxml/minimal.docx");

        let parsed = parser.parse_bytes(&bytes).expect("parse bytes");
        assert_eq!(parsed.format(), DocumentFormat::WordProcessing);
        assert!(parsed.document().is_some());

        let mut store = IrStore::new();
        parser
            .scan_security_bytes(&bytes, &mut store)
            .expect("scan security bytes");
    }

    #[test]
    fn parser_port_reader_and_file_with_bytes_preserve_payload() {
        let parser = DocumentParser::new();
        let bytes = fixture_bytes("fixtures/ooxml/minimal.xlsx");

        let parsed =
            ParserPort::parse_reader(&parser, Cursor::new(bytes.clone())).expect("parse reader");
        assert_eq!(parsed.format(), DocumentFormat::Spreadsheet);

        let (parsed_with_reader_bytes, reader_payload) =
            ParserPort::parse_reader_with_bytes(&parser, Cursor::new(bytes.clone()))
                .expect("parse reader with bytes");
        assert_eq!(
            parsed_with_reader_bytes.format(),
            DocumentFormat::Spreadsheet
        );
        assert_eq!(reader_payload, bytes);

        let ts = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system time")
            .as_nanos();
        let tmp_path = std::env::temp_dir().join(format!("docir-app-adapters-{ts}.xlsx"));
        fs::write(&tmp_path, &reader_payload).expect("write temp file");

        let (parsed_with_file_bytes, file_payload) =
            ParserPort::parse_file_with_bytes(&parser, &tmp_path).expect("parse file with bytes");
        assert_eq!(parsed_with_file_bytes.format(), DocumentFormat::Spreadsheet);
        assert_eq!(file_payload, reader_payload);
        fs::remove_file(tmp_path).expect("remove temp file");
    }

    #[test]
    fn app_parser_new_uses_given_parser_and_config() {
        let parser = AppParser::new(DocumentParser::new(), ParserConfig::default());
        let bytes = fixture_bytes("fixtures/ooxml/minimal.pptx");

        let parsed = parser
            .parse_reader(Cursor::new(bytes))
            .expect("parse reader");
        assert_eq!(parsed.format(), DocumentFormat::Presentation);
    }
}
