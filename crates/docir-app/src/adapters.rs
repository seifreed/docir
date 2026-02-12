//! Infrastructure adapters for application ports.

use crate::{AppResult, ParsedDocument, ParserConfig, ParserPort, SecurityScannerPort};
use docir_core::visitor::IrStore;
use docir_parser::zip_handler::SecureZipReader;
use docir_parser::{DefaultSecurityScanner, DocumentParser, ParseError, SecurityScanner};
use std::io::{Cursor, Read, Seek};
use std::path::Path;

/// Parser adapter that bundles a configured parser with its config.
pub struct AppParser {
    parser: DocumentParser,
    config: docir_parser::ParserConfig,
}

impl AppParser {
    pub fn new(parser: DocumentParser, config: ParserConfig) -> Self {
        let parser_config = config.to_parser_config();
        Self {
            parser,
            config: parser_config,
        }
    }

    pub fn with_config(config: ParserConfig) -> Self {
        let parser_config = config.to_parser_config();
        let parser = DocumentParser::with_config(parser_config.clone());
        Self {
            parser,
            config: parser_config,
        }
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

fn scan_security_bytes(
    config: &docir_parser::ParserConfig,
    data: &[u8],
    store: &mut IrStore,
) -> Result<(), ParseError> {
    if !is_zip_container(data) {
        return Ok(());
    }
    let mut zip = SecureZipReader::new(Cursor::new(data), config.zip_config.clone())?;
    let scanner = DefaultSecurityScanner;
    scanner.scan_ooxml(config, &mut zip, store)
}

fn is_zip_container(data: &[u8]) -> bool {
    data.len() >= 4 && data[0] == b'P' && data[1] == b'K'
}
