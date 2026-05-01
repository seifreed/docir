use crate::parser::{ParsedDocument, ParserConfig};
use crate::{
    HwpParser, HwpxParser, LegacyOfficeParser, OdfParser, OoxmlParser, ParseError, RtfParser,
};
use std::io::{Read, Seek};

pub(super) enum DetectedFormat {
    Ooxml,
    Odf,
    Hwpx,
    Hwp,
    LegacyOffice,
    Rtf,
}

pub(super) enum ParserDispatch {
    Ooxml(OoxmlParser),
    Odf(OdfParser),
    Hwpx(HwpxParser),
    Hwp(HwpParser),
    LegacyOffice(LegacyOfficeParser),
    Rtf(RtfParser),
}

impl ParserDispatch {
    pub(super) fn parse_reader<R: Read + Seek>(
        self,
        reader: R,
    ) -> Result<ParsedDocument, ParseError> {
        match self {
            ParserDispatch::Ooxml(parser) => parser.parse_reader(reader),
            ParserDispatch::Odf(parser) => parser.parse_reader(reader),
            ParserDispatch::Hwpx(parser) => parser.parse_reader(reader),
            ParserDispatch::Hwp(parser) => parser.parse_reader(reader),
            ParserDispatch::LegacyOffice(parser) => parser.parse_reader(reader),
            ParserDispatch::Rtf(parser) => parser.parse_reader(reader),
        }
    }
}

pub(super) fn build_parser(format: DetectedFormat, config: ParserConfig) -> ParserDispatch {
    match format {
        DetectedFormat::Ooxml => ParserDispatch::Ooxml(OoxmlParser::with_config(config)),
        DetectedFormat::Odf => ParserDispatch::Odf(OdfParser::with_config(config)),
        DetectedFormat::Hwpx => ParserDispatch::Hwpx(HwpxParser::with_config(config)),
        DetectedFormat::Hwp => ParserDispatch::Hwp(HwpParser::with_config(config)),
        DetectedFormat::LegacyOffice => {
            ParserDispatch::LegacyOffice(LegacyOfficeParser::with_config(config))
        }
        DetectedFormat::Rtf => ParserDispatch::Rtf(RtfParser::with_config(config)),
    }
}
