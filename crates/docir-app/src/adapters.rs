//! Infrastructure adapters for application ports.

use crate::{AppResult, ParsedDocument, ParserPort, SecurityScannerPort};
use docir_core::visitor::IrStore;
use docir_parser::{scan_security_bytes, DocumentParser, ParserConfig};
use std::io::{Read, Seek};
use std::path::Path;

/// Parser adapter that bundles a configured parser with its config.
pub struct AppParser {
    parser: DocumentParser,
    config: ParserConfig,
}

impl AppParser {
    pub fn new(parser: DocumentParser, config: ParserConfig) -> Self {
        Self { parser, config }
    }

    pub fn with_config(config: ParserConfig) -> Self {
        let parser = DocumentParser::with_config(config.clone());
        Self::new(parser, config)
    }
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
