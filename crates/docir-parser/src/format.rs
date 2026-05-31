use crate::ParseError;
use crate::parser::ParsedDocument;
use std::io::{Read, Seek};

pub trait FormatParser {
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError>;
}
