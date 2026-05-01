//! Minimal legacy Office CFB parser for `.doc`, `.xls`, and `.ppt`.

use crate::error::ParseError;
use crate::format::FormatParser;
use crate::input::read_all_with_limit;
use crate::ole::Cfb;
use crate::parser::{ParseMetrics, ParsedDocument, ParserConfig};
use docir_core::ir::{Document, ExtensionPart, ExtensionPartKind, IRNode};
use docir_core::types::{DocumentFormat, SourceSpan};
use docir_core::visitor::IrStore;
use std::io::{Read, Seek};

/// Minimal parser for legacy Office documents stored in CFB/OLE containers.
pub struct LegacyOfficeParser {
    config: ParserConfig,
}

impl FormatParser for LegacyOfficeParser {
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        self.parse_reader(reader)
    }
}

impl Default for LegacyOfficeParser {
    fn default() -> Self {
        Self::new()
    }
}

impl LegacyOfficeParser {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            config: ParserConfig::default(),
        }
    }

    /// Public API entrypoint: with_config.
    pub fn with_config(config: ParserConfig) -> Self {
        Self { config }
    }

    crate::impl_parse_entrypoints!();

    /// Parses legacy Office CFB content into a minimal IR.
    pub fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        let data = read_all_with_limit(reader, self.config.max_input_size)?;
        let cfb = Cfb::parse(data)?;
        let format = probe_legacy_office_format(&cfb).ok_or_else(|| {
            ParseError::UnsupportedFormat("Unsupported CFB/OLE Office container".to_string())
        })?;

        let mut store = IrStore::new();
        let mut document = Document::new(format);
        document.span = Some(SourceSpan::new("cfb:/"));

        let mut stream_names = cfb.list_streams();
        stream_names.sort();
        for path in stream_names {
            let size = cfb.stream_size(&path).unwrap_or(0);
            let mut part = ExtensionPart::new(&path, size, ExtensionPartKind::Legacy);
            part.span = Some(SourceSpan::new(&path));
            let id = part.id;
            store.insert(IRNode::ExtensionPart(part));
            document.shared_parts.push(id);
        }

        let root_id = document.id;
        store.insert(IRNode::Document(document));
        Ok(ParsedDocument {
            root_id,
            format,
            store,
            metrics: Some(ParseMetrics::default()),
        })
    }
}

/// Detects whether a CFB container looks like a legacy Office document.
pub fn probe_legacy_office_format(cfb: &Cfb) -> Option<DocumentFormat> {
    let streams = cfb.list_streams();
    let has = |needle: &str| {
        streams
            .iter()
            .any(|stream| stream.eq_ignore_ascii_case(needle))
    };

    if has("WordDocument") {
        return Some(DocumentFormat::WordProcessing);
    }
    if has("Workbook") || has("Book") {
        return Some(DocumentFormat::Spreadsheet);
    }
    if has("PowerPoint Document") || has("Current User") {
        return Some(DocumentFormat::Presentation);
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::DocumentParser;
    use crate::test_support::build_test_cfb;

    #[test]
    fn probe_legacy_format_detects_word_stream() {
        let bytes = build_test_cfb(&[("WordDocument", b"doc")]);
        let cfb = Cfb::parse(bytes).expect("cfb");
        assert_eq!(
            probe_legacy_office_format(&cfb),
            Some(DocumentFormat::WordProcessing)
        );
    }

    #[test]
    fn parser_builds_legacy_shared_parts() {
        let bytes = build_test_cfb(&[("WordDocument", b"doc"), ("1Table", b"tbl")]);
        let parsed = LegacyOfficeParser::new()
            .parse_bytes(&bytes)
            .expect("parse");
        let doc = parsed.document().expect("document");
        assert_eq!(parsed.format, DocumentFormat::WordProcessing);
        assert_eq!(doc.shared_parts.len(), 2);
        assert_eq!(
            doc.span.as_ref().map(|span| span.file_path.as_str()),
            Some("cfb:/")
        );
    }

    #[test]
    fn document_parser_dispatches_cfb_to_legacy_office_parser() {
        let bytes = build_test_cfb(&[("Workbook", b"wb")]);
        let parsed = DocumentParser::new()
            .parse_bytes(&bytes)
            .expect("dispatch parse");
        assert_eq!(parsed.format, DocumentFormat::Spreadsheet);
        assert_eq!(
            parsed
                .document()
                .and_then(|doc| doc.span.as_ref())
                .map(|span| span.file_path.as_str()),
            Some("cfb:/")
        );
    }
}
