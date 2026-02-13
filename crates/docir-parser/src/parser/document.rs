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

pub(super) fn parse_activex_xml(
    xml: &str,
    _path: &str,
) -> Option<docir_core::security::ActiveXControl> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut control = docir_core::security::ActiveXControl::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    match key {
                        b"name" => control.name = Some(val.clone()),
                        b"clsid" | b"classid" => control.clsid = Some(val.clone()),
                        b"progid" => control.prog_id = Some(val.clone()),
                        _ => {
                            let k = String::from_utf8_lossy(key).to_string();
                            control.properties.push((k, val));
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    if control.name.is_some() || control.clsid.is_some() || control.prog_id.is_some() {
        Some(control)
    } else {
        None
    }
}

pub(super) fn map_calamine_error(err: calamine::CellErrorType) -> CellError {
    use calamine::CellErrorType::*;
    match err {
        Null => CellError::Null,
        Div0 => CellError::DivZero,
        Value => CellError::Value,
        Ref => CellError::Ref,
        Name => CellError::Name,
        Num => CellError::Num,
        NA => CellError::NA,
        GettingData => CellError::GettingData,
    }
}

pub(super) fn parse_smartart_part(
    xml: &str,
    path: &str,
) -> Result<docir_core::ir::SmartArtPart, ParseError> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut root_element: Option<String> = None;
    let mut point_count: u32 = 0;
    let mut connection_count: u32 = 0;
    let mut rel_ids: Vec<String> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if root_element.is_none() {
                    root_element = Some(String::from_utf8_lossy(e.name().as_ref()).to_string());
                }
                let name_buf = e.name().as_ref().to_vec();
                let name = name_buf.as_slice();
                if name.ends_with(b":pt") || name == b"dgm:pt" {
                    point_count += 1;
                }
                if name.ends_with(b":cxn") || name == b"dgm:cxn" {
                    connection_count += 1;
                }
                if name.ends_with(b":relIds") || name == b"dgm:relIds" {
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        if key == b"r:dm" || key == b"r:lo" || key == b"r:qs" || key == b"r:cs" {
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if !val.is_empty() {
                                rel_ids.push(val);
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(docir_core::ir::SmartArtPart {
        id: NodeId::new(),
        kind: "diagram".to_string(),
        path: path.to_string(),
        root_element,
        point_count: if point_count > 0 {
            Some(point_count)
        } else {
            None
        },
        connection_count: if connection_count > 0 {
            Some(connection_count)
        } else {
            None
        },
        rel_ids,
        span: Some(SourceSpan::new(path)),
    })
}

pub(super) fn parse_chart_data(
    xml: &str,
    chart_path: &str,
    store: &mut IrStore,
) -> Result<NodeId, ParseError> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut chart = docir_core::ir::ChartData::new();
    chart.span = Some(SourceSpan::new(chart_path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_title = false;
    let mut in_series = false;
    let mut section: Option<&[u8]> = None;
    let mut current_series: Option<docir_core::ir::ChartSeries> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name_bytes = e.name().as_ref().to_vec();
                let name = local_name(&name_bytes);
                if name.ends_with(b"Chart") || name.ends_with(b"chart") {
                    chart.chart_type = Some(String::from_utf8_lossy(name).to_string());
                }
                if name == b"ser" {
                    in_series = true;
                    current_series = Some(docir_core::ir::ChartSeries::new());
                }
                if !in_series && (name == b"title") {
                    in_title = true;
                }
                if in_series {
                    if name == b"tx" {
                        section = Some(b"tx");
                    } else if name == b"cat" {
                        section = Some(b"cat");
                    } else if name == b"val" {
                        section = Some(b"val");
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name_bytes = e.name().as_ref().to_vec();
                let name = local_name(&name_bytes);
                if name == b"title" {
                    in_title = false;
                }
                if name == b"ser" {
                    in_series = false;
                    section = None;
                    if let Some(series) = current_series.take() {
                        if let Some(name) = &series.name {
                            chart.series.push(name.clone());
                        }
                        chart.series_data.push(series);
                    }
                }
                if name == b"tx" || name == b"cat" || name == b"val" {
                    section = None;
                }
            }
            Ok(Event::Text(e)) => {
                let text = e.unescape().unwrap_or_default().to_string();
                if in_title && chart.title.is_none() {
                    chart.title = Some(text);
                } else if in_series {
                    let trimmed = text.trim();
                    if trimmed.is_empty() {
                        // skip
                    } else if let Some(series) = current_series.as_mut() {
                        match section {
                            Some(b"tx") => {
                                if series.name.is_none() {
                                    series.name = Some(trimmed.to_string());
                                }
                            }
                            Some(b"cat") => {
                                series.categories.push(trimmed.to_string());
                            }
                            Some(b"val") => {
                                series.values.push(trimmed.to_string());
                            }
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: chart_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    let id = chart.id;
    store.insert(IRNode::ChartData(chart));
    Ok(id)
}

/// Convert a .rels path to its source part path when possible.

impl Default for OoxmlParser {
    fn default() -> Self {
        Self::new()
    }
}

/// Helper function to encode bytes as hex.
pub(super) mod hex {
    pub fn encode(data: impl AsRef<[u8]>) -> String {
        data.as_ref().iter().map(|b| format!("{:02x}", b)).collect()
    }
}
