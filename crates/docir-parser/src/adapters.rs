use crate::format::FormatParser;
use crate::{HwpParser, HwpxParser, OdfParser, OoxmlParser, ParseError, ParserConfig, RtfParser};
use std::io::{Read, Seek};

macro_rules! define_adapter {
    ($adapter:ident, $parser:ty) => {
        pub struct $adapter {
            config: ParserConfig,
        }

        impl $adapter {
            pub fn new(config: ParserConfig) -> Self {
                Self { config }
            }
        }

        impl FormatParser for $adapter {
            fn parse_reader<R: Read + Seek>(
                &self,
                reader: R,
            ) -> Result<crate::parser::ParsedDocument, ParseError> {
                <$parser>::with_config(self.config.clone()).parse_reader(reader)
            }
        }
    };
}

define_adapter!(OoxmlAdapter, OoxmlParser);
define_adapter!(OdfAdapter, OdfParser);
define_adapter!(HwpxAdapter, HwpxParser);
define_adapter!(HwpAdapter, HwpParser);
define_adapter!(RtfAdapter, RtfParser);
