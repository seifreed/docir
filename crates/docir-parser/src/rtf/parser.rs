//! RTF parser entrypoints.

use crate::error::ParseError;
use crate::format::FormatParser;
use crate::input::read_all_with_limit;
use crate::parse_utils::finalize_and_normalize;
use crate::parser::{ParsedDocument, ParserConfig};
use docir_core::ir::IRNode;
use docir_core::types::DocumentFormat;
use docir_core::visitor::IrStore;
use std::io::{Read, Seek};

use super::{is_rtf_bytes, parse_rtf, RtfCursor, RtfParseContext};

/// Parser for RTF documents.
pub struct RtfParser {
    config: ParserConfig,
}

impl FormatParser for RtfParser {
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        self.parse_reader(reader)
    }
}

impl RtfParser {
    /// Creates a new parser with default configuration.
    pub fn new() -> Self {
        Self {
            config: ParserConfig::default(),
        }
    }

    /// Creates a new parser with custom configuration.
    pub fn with_config(config: ParserConfig) -> Self {
        Self { config }
    }

    crate::impl_parse_entrypoints!();

    /// Parses from any reader.
    pub fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        let data = read_all_with_limit(reader, self.config.max_input_size)?;
        if !is_rtf_bytes(&data) {
            return Err(ParseError::UnsupportedFormat(
                "Missing RTF header".to_string(),
            ));
        }

        let mut store = IrStore::new();
        let mut doc = docir_core::ir::Document::new(DocumentFormat::Rtf);

        let mut ctx = RtfParseContext::new(
            self.config.rtf.max_group_depth,
            self.config.rtf.max_object_hex_len,
        );
        let mut cursor = RtfCursor::new(&data);
        parse_rtf(&mut cursor, &mut ctx, &mut store)?;

        if let Some(style_set) = ctx.style_set.take() {
            let style_id = style_set.id;
            store.insert(IRNode::StyleSet(style_set));
            doc.styles = Some(style_id);
        }

        for section in ctx.sections {
            doc.content.push(section);
        }
        for media in ctx.media_assets {
            doc.shared_parts.push(media);
        }

        Ok(finalize_and_normalize(DocumentFormat::Rtf, store, doc))
    }
}
