use crate::format::FormatParser;
use crate::{HwpParser, HwpxParser, OdfParser, OoxmlParser, ParseError, ParserConfig, RtfParser};
use std::io::{Read, Seek};

pub struct OoxmlAdapter {
    config: ParserConfig,
}

impl OoxmlAdapter {
    pub fn new(config: ParserConfig) -> Self {
        Self { config }
    }
}

impl FormatParser for OoxmlAdapter {
    fn parse_reader<R: Read + Seek>(
        &self,
        reader: R,
    ) -> Result<crate::parser::ParsedDocument, ParseError> {
        OoxmlParser::with_config(self.config.clone()).parse_reader(reader)
    }
}

pub struct OdfAdapter {
    config: ParserConfig,
}

impl OdfAdapter {
    pub fn new(config: ParserConfig) -> Self {
        Self { config }
    }
}

impl FormatParser for OdfAdapter {
    fn parse_reader<R: Read + Seek>(
        &self,
        reader: R,
    ) -> Result<crate::parser::ParsedDocument, ParseError> {
        OdfParser::with_config(self.config.clone()).parse_reader(reader)
    }
}

pub struct HwpxAdapter {
    config: ParserConfig,
}

impl HwpxAdapter {
    pub fn new(config: ParserConfig) -> Self {
        Self { config }
    }
}

impl FormatParser for HwpxAdapter {
    fn parse_reader<R: Read + Seek>(
        &self,
        reader: R,
    ) -> Result<crate::parser::ParsedDocument, ParseError> {
        HwpxParser::with_config(self.config.clone()).parse_reader(reader)
    }
}

pub struct HwpAdapter {
    config: ParserConfig,
}

impl HwpAdapter {
    pub fn new(config: ParserConfig) -> Self {
        Self { config }
    }
}

impl FormatParser for HwpAdapter {
    fn parse_reader<R: Read + Seek>(
        &self,
        reader: R,
    ) -> Result<crate::parser::ParsedDocument, ParseError> {
        HwpParser::with_config(self.config.clone()).parse_reader(reader)
    }
}

pub struct RtfAdapter {
    config: ParserConfig,
}

impl RtfAdapter {
    pub fn new(config: ParserConfig) -> Self {
        Self { config }
    }
}

impl FormatParser for RtfAdapter {
    fn parse_reader<R: Read + Seek>(
        &self,
        reader: R,
    ) -> Result<crate::parser::ParsedDocument, ParseError> {
        RtfParser::with_config(self.config.clone()).parse_reader(reader)
    }
}
