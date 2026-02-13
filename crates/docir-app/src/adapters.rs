//! Infrastructure adapters for application ports.

use crate::{
    AppParseError, AppResult, ParsedDocument, ParserConfig, ParserPort, SecurityScannerPort,
};
use docir_core::visitor::IrStore;
use docir_parser::parser::ParsedDocument as ParserParsedDocument;
use docir_parser::{scan_security_bytes as scan_parser_bytes, DocumentParser, ParseError};
use std::io::{Read, Seek};
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
        wrap_parsed(self.parser.parse_file(path))
    }

    fn parse_bytes(&self, data: &[u8]) -> AppResult<ParsedDocument> {
        wrap_parsed(self.parser.parse_bytes(data))
    }

    fn parse_reader<R: Read + Seek>(&self, reader: R) -> AppResult<ParsedDocument> {
        wrap_parsed(self.parser.parse_reader(reader))
    }

    fn parse_file_with_bytes<P: AsRef<Path>>(
        &self,
        path: P,
    ) -> AppResult<(ParsedDocument, Vec<u8>)> {
        wrap_parsed_with_bytes(self.parser.parse_file_with_bytes(path))
    }

    fn parse_reader_with_bytes<R: Read + Seek>(
        &self,
        reader: R,
    ) -> AppResult<(ParsedDocument, Vec<u8>)> {
        wrap_parsed_with_bytes(self.parser.parse_reader_with_bytes(reader))
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
    result
        .map(ParsedDocument::new)
        .map_err(AppParseError::from)
        .map_err(Into::into)
}

fn wrap_parsed_with_bytes(
    result: Result<(ParserParsedDocument, Vec<u8>), ParseError>,
) -> AppResult<(ParsedDocument, Vec<u8>)> {
    result
        .map(|(parsed, data)| (ParsedDocument::new(parsed), data))
        .map_err(AppParseError::from)
        .map_err(Into::into)
}
