use super::*;
use crate::parse_utils::is_zip_container;

/// Unified parser that dispatches OOXML vs ODF based on package signature.
pub struct DocumentParser {
    config: ParserConfig,
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
    pub fn parse_reader<R: Read + Seek>(
        &self,
        mut reader: R,
    ) -> Result<ParsedDocument, ParseError> {
        enforce_input_size(&mut reader, self.config.max_input_size)?;
        let detected = self.detect_format(&mut reader)?;
        self.validate_detected(&detected, &mut reader)?;
        let parsed = self.parse_detected(detected, reader)?;
        self.normalize_parsed(parsed)
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
            return Ok(formats::DetectedFormat::Hwp);
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

    fn validate_detected<R: Read + Seek>(
        &self,
        _detected: &formats::DetectedFormat,
        _reader: &mut R,
    ) -> Result<(), ParseError> {
        Ok(())
    }

    fn parse_detected<R: Read + Seek>(
        &self,
        detected: formats::DetectedFormat,
        reader: R,
    ) -> Result<ParsedDocument, ParseError> {
        formats::build_parser(detected, self.config.clone()).parse_reader(reader)
    }

    fn normalize_parsed(&self, parsed: ParsedDocument) -> Result<ParsedDocument, ParseError> {
        Ok(parsed)
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

impl Default for OoxmlParser {
    fn default() -> Self {
        Self::new()
    }
}
