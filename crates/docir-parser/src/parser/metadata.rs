use super::{OoxmlParser, ParseError};
use crate::zip_handler::SecureZipReader;
use docir_core::ir::{CustomProperty, DocumentMetadata, PropertyValue};
use docir_core::types::NodeId;
use std::io::{Read, Seek};

impl OoxmlParser {
    /// Parse document metadata.
    pub(super) fn parse_metadata<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
    ) -> Result<Option<NodeId>, ParseError> {
        if zip.contains("docProps/core.xml") {
            let metadata = self.build_metadata(zip);
            Ok(metadata.map(|m| m.id))
        } else {
            Ok(None)
        }
    }

    /// Build metadata from core.xml and app.xml.
    pub(super) fn build_metadata<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
    ) -> Option<DocumentMetadata> {
        let mut metadata = DocumentMetadata::new();

        // Parse core.xml (Dublin Core properties)
        if let Ok(core_xml) = zip.read_file_string("docProps/core.xml") {
            self.parse_core_properties(&core_xml, &mut metadata);
        }

        // Parse app.xml (application properties)
        if let Ok(app_xml) = zip.read_file_string("docProps/app.xml") {
            self.parse_app_properties(&app_xml, &mut metadata);
        }

        // Parse custom.xml (custom properties)
        if let Ok(custom_xml) = zip.read_file_string("docProps/custom.xml") {
            self.parse_custom_properties(&custom_xml, &mut metadata);
        }

        Some(metadata)
    }

    /// Parse core.xml properties.
    fn parse_core_properties(&self, xml: &str, metadata: &mut DocumentMetadata) {
        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_element = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    current_element = String::from_utf8_lossy(e.name().as_ref()).to_string();
                }
                Ok(Event::Text(e)) => {
                    let text = e.unescape().unwrap_or_default().to_string();
                    match current_element.as_str() {
                        "dc:title" | "title" => metadata.title = Some(text),
                        "dc:subject" | "subject" => metadata.subject = Some(text),
                        "dc:creator" | "creator" => metadata.creator = Some(text),
                        "cp:keywords" | "keywords" => metadata.keywords = Some(text),
                        "dc:description" | "description" => metadata.description = Some(text),
                        "cp:lastModifiedBy" | "lastModifiedBy" => {
                            metadata.last_modified_by = Some(text)
                        }
                        "cp:revision" | "revision" => metadata.revision = Some(text),
                        "dcterms:created" | "created" => metadata.created = Some(text),
                        "dcterms:modified" | "modified" => metadata.modified = Some(text),
                        "cp:category" | "category" => metadata.category = Some(text),
                        "cp:contentStatus" | "contentStatus" => {
                            metadata.content_status = Some(text)
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(_)) => {
                    current_element.clear();
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }
    }

    /// Parse app.xml properties.
    fn parse_app_properties(&self, xml: &str, metadata: &mut DocumentMetadata) {
        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_element = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    current_element = String::from_utf8_lossy(e.name().as_ref()).to_string();
                }
                Ok(Event::Text(e)) => {
                    let text = e.unescape().unwrap_or_default().to_string();
                    match current_element.as_str() {
                        "Application" => metadata.application = Some(text),
                        "AppVersion" => metadata.app_version = Some(text),
                        "Company" => metadata.company = Some(text),
                        "Manager" => metadata.manager = Some(text),
                        _ => {}
                    }
                }
                Ok(Event::End(_)) => {
                    current_element.clear();
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }
    }

    /// Parse custom properties from custom.xml.
    pub(super) fn parse_custom_properties(&self, xml: &str, metadata: &mut DocumentMetadata) {
        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_prop: Option<CustomProperty> = None;
        let mut current_value_tag: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                    if name.ends_with("property") {
                        let mut prop = CustomProperty {
                            name: String::new(),
                            value: PropertyValue::String(String::new()),
                            format_id: None,
                            property_id: None,
                        };
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"name" => {
                                    prop.name = String::from_utf8_lossy(&attr.value).to_string()
                                }
                                b"fmtid" => {
                                    prop.format_id =
                                        Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"pid" => {
                                    prop.property_id =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                                }
                                _ => {}
                            }
                        }
                        current_prop = Some(prop);
                    } else if name.starts_with("vt:") || name.contains(':') {
                        current_value_tag = Some(name);
                    }
                }
                Ok(Event::Text(e)) => {
                    if let Some(tag) = &current_value_tag {
                        if let Some(prop) = current_prop.as_mut() {
                            let text = e.unescape().unwrap_or_default().to_string();
                            prop.value = match tag.as_str() {
                                "vt:lpwstr" | "vt:lpstr" | "vt:bstr" => PropertyValue::String(text),
                                "vt:i2" | "vt:i4" | "vt:int" | "vt:integer" => {
                                    PropertyValue::Integer(text.parse::<i64>().unwrap_or(0))
                                }
                                "vt:r4" | "vt:r8" | "vt:float" => {
                                    PropertyValue::Float(text.parse::<f64>().unwrap_or(0.0))
                                }
                                "vt:bool" => PropertyValue::Boolean(text == "true" || text == "1"),
                                "vt:filetime" => PropertyValue::DateTime(text),
                                "vt:blob" => PropertyValue::Blob(text),
                                _ => PropertyValue::String(text),
                            };
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                    if name.ends_with("property") {
                        if let Some(prop) = current_prop.take() {
                            metadata.custom_properties.push(prop);
                        }
                    } else if current_value_tag.as_deref() == Some(&name) {
                        current_value_tag = None;
                    }
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }
    }
}
