use super::{
    enforce_input_size, formats, is_hwpx_mimetype, is_ole_container, is_rtf_bytes,
    read_all_with_limit, run_parser_pipeline, NormalizeStage, OoxmlParser, ParseError, ParseStage,
    ParsedDocument, ParserConfig, PostprocessStage, Read, SecureZipReader, Seek, SeekFrom,
};
use crate::legacy_office::probe_legacy_office_format;
use crate::parse_utils::is_zip_container;
use std::path::Path;

/// Unified parser that dispatches OOXML vs ODF based on package signature.
pub struct DocumentParser {
    config: ParserConfig,
}

impl Default for DocumentParser {
    fn default() -> Self {
        Self::new()
    }
}

impl DocumentParser {
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
        run_parser_pipeline(self, reader)
    }

    fn detect_format<R: Read + Seek>(
        &self,
        reader: &mut R,
    ) -> Result<formats::DetectedFormat, ParseError> {
        let mut probe = [0u8; 16];
        let read = reader.read(&mut probe)?;
        reader.seek(SeekFrom::Start(0))?;
        let head = &probe[..read];

        if is_rtf_bytes(head) {
            return Ok(formats::DetectedFormat::Rtf);
        }

        if is_ole_container(head) {
            let data = read_all_with_limit(&mut *reader, self.config.max_input_size)?;
            let cfb = crate::ole::Cfb::parse(data)?;
            reader.seek(SeekFrom::Start(0))?;
            if cfb.has_stream("FileHeader") {
                return Ok(formats::DetectedFormat::Hwp);
            }
            if probe_legacy_office_format(&cfb).is_some() {
                return Ok(formats::DetectedFormat::LegacyOffice);
            }
            return Err(ParseError::UnsupportedFormat(
                "Unknown OLE/CFB container (not HWP or legacy Office)".to_string(),
            ));
        }

        if !is_zip_container(head) {
            return Err(ParseError::UnsupportedFormat(
                "Unknown package format (not OLE/CFB or ZIP)".to_string(),
            ));
        }

        let mut inspector = SecureZipReader::new(&mut *reader, self.config.zip_config.clone())?;
        let is_ooxml = inspector.contains("[Content_Types].xml");
        let is_odf = inspector.contains("mimetype");
        let is_hwpx = if is_odf {
            inspector
                .read_file_string("mimetype")
                .map(|m| is_hwpx_mimetype(&m))
                .unwrap_or(false)
        } else {
            false
        };
        drop(inspector);
        reader.seek(SeekFrom::Start(0))?;

        if is_ooxml {
            return Ok(formats::DetectedFormat::Ooxml);
        }
        if is_hwpx {
            return Ok(formats::DetectedFormat::Hwpx);
        }
        if is_odf {
            return Ok(formats::DetectedFormat::Odf);
        }

        Err(ParseError::UnsupportedFormat(
            "Unknown package format (missing [Content_Types].xml and mimetype)".to_string(),
        ))
    }

    fn parse_detected<R: Read + Seek>(
        &self,
        detected: formats::DetectedFormat,
        reader: R,
    ) -> Result<ParsedDocument, ParseError> {
        formats::build_parser(detected, self.config.clone()).parse_reader(reader)
    }

    /// Parses from a file and returns parsed document with raw bytes.
    pub fn parse_file_with_bytes<P: AsRef<Path>>(
        &self,
        path: P,
    ) -> Result<(ParsedDocument, Vec<u8>), ParseError> {
        let reader = crate::input::open_reader(path)?;
        let data = read_all_with_limit(reader, self.config.max_input_size)?;
        let parsed = self.parse_bytes(&data)?;
        Ok((parsed, data))
    }

    /// Parses from a reader and returns parsed document with raw bytes.
    pub fn parse_reader_with_bytes<R: Read + Seek>(
        &self,
        reader: R,
    ) -> Result<(ParsedDocument, Vec<u8>), ParseError> {
        let data = read_all_with_limit(reader, self.config.max_input_size)?;
        let parsed = self.parse_bytes(&data)?;
        Ok((parsed, data))
    }
}

impl ParseStage for DocumentParser {
    fn parse_stage<R: Read + Seek>(&self, mut reader: R) -> Result<ParsedDocument, ParseError> {
        enforce_input_size(&mut reader, self.config.max_input_size)?;
        let detected = self.detect_format(&mut reader)?;
        self.parse_detected(detected, reader)
    }
}

impl NormalizeStage for DocumentParser {}

impl PostprocessStage for DocumentParser {}

impl Default for OoxmlParser {
    fn default() -> Self {
        Self::new()
    }
}
