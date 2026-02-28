use super::{OoxmlParser, ParseError};
use crate::xml_utils::reader_from_str;
use crate::zip_handler::PackageReader;
use docir_core::ir::{CustomProperty, DocumentMetadata, PropertyValue};
use docir_core::types::NodeId;

impl OoxmlParser {
    /// Parse document metadata.
    pub(super) fn parse_metadata(
        &self,
        zip: &mut impl PackageReader,
    ) -> Result<Option<NodeId>, ParseError> {
        if zip.contains("docProps/core.xml") {
            let metadata = self.build_metadata(zip);
            Ok(metadata.map(|m| m.id))
        } else {
            Ok(None)
        }
    }

    /// Build metadata from core.xml and app.xml.
    pub(super) fn build_metadata(&self, zip: &mut impl PackageReader) -> Option<DocumentMetadata> {
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

        let mut reader = reader_from_str(xml);

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

        let mut reader = reader_from_str(xml);

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

        let mut reader = reader_from_str(xml);

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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::zip_handler::PackageReader;
    use docir_core::ir::PropertyValue;
    use std::collections::HashMap;

    struct TestPackageReader {
        files: HashMap<String, Vec<u8>>,
    }

    impl TestPackageReader {
        fn new(entries: &[(&str, &str)]) -> Self {
            let files = entries
                .iter()
                .map(|(path, body)| ((*path).to_string(), body.as_bytes().to_vec()))
                .collect();
            Self { files }
        }
    }

    impl PackageReader for TestPackageReader {
        fn contains(&self, name: &str) -> bool {
            self.files.contains_key(name)
        }

        fn read_file(&mut self, name: &str) -> Result<Vec<u8>, ParseError> {
            self.files
                .get(name)
                .cloned()
                .ok_or_else(|| ParseError::MissingPart(name.to_string()))
        }

        fn read_file_string(&mut self, name: &str) -> Result<String, ParseError> {
            let bytes = self.read_file(name)?;
            String::from_utf8(bytes)
                .map_err(|e| ParseError::Encoding(format!("Invalid UTF-8 in {name}: {e}")))
        }

        fn file_size(&mut self, name: &str) -> Result<u64, ParseError> {
            self.files
                .get(name)
                .map(|v| v.len() as u64)
                .ok_or_else(|| ParseError::MissingPart(name.to_string()))
        }

        fn file_names(&self) -> Vec<String> {
            self.files.keys().cloned().collect()
        }

        fn list_prefix(&self, prefix: &str) -> Vec<String> {
            self.files
                .keys()
                .filter(|name| name.starts_with(prefix))
                .cloned()
                .collect()
        }

        fn list_suffix(&self, suffix: &str) -> Vec<String> {
            self.files
                .keys()
                .filter(|name| name.ends_with(suffix))
                .cloned()
                .collect()
        }
    }

    #[test]
    fn parse_metadata_returns_none_without_core_properties_part() {
        let parser = OoxmlParser::new();
        let mut zip = TestPackageReader::new(&[]);
        let id = parser.parse_metadata(&mut zip).expect("metadata parse succeeds");
        assert!(id.is_none());
    }

    #[test]
    fn build_metadata_parses_core_app_and_custom_typed_properties() {
        let core = r#"
            <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                               xmlns:dc="http://purl.org/dc/elements/1.1/"
                               xmlns:dcterms="http://purl.org/dc/terms/">
              <dc:title>Threat Report</dc:title>
              <dc:subject>Forensics</dc:subject>
              <dc:creator>analyst</dc:creator>
              <cp:keywords>malware,dde</cp:keywords>
              <dc:description>summary</dc:description>
              <cp:lastModifiedBy>reviewer</cp:lastModifiedBy>
              <cp:revision>7</cp:revision>
              <dcterms:created>2025-05-01T10:00:00Z</dcterms:created>
              <dcterms:modified>2025-05-01T11:00:00Z</dcterms:modified>
              <cp:category>incident</cp:category>
              <cp:contentStatus>final</cp:contentStatus>
            </cp:coreProperties>
        "#;
        let app = r#"
            <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
              <Application>docir</Application>
              <AppVersion>1.0</AppVersion>
              <Company>ACME</Company>
              <Manager>SecOps</Manager>
            </Properties>
        "#;
        let custom = r#"
            <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
                        xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
              <property fmtid="{A}" pid="2" name="Build"><vt:i4>42</vt:i4></property>
              <property fmtid="{A}" pid="3" name="Ratio"><vt:r8>3.25</vt:r8></property>
              <property fmtid="{A}" pid="4" name="Flag"><vt:bool>1</vt:bool></property>
              <property fmtid="{A}" pid="5" name="SeenAt"><vt:filetime>2025-05-01T00:00:00Z</vt:filetime></property>
              <property fmtid="{A}" pid="6" name="Payload"><vt:blob>QUJD</vt:blob></property>
              <property fmtid="{A}" pid="7" name="Tag"><vt:lpwstr>alert</vt:lpwstr></property>
            </Properties>
        "#;

        let parser = OoxmlParser::new();
        let mut zip = TestPackageReader::new(&[
            ("docProps/core.xml", core),
            ("docProps/app.xml", app),
            ("docProps/custom.xml", custom),
        ]);

        let metadata = parser.build_metadata(&mut zip).expect("metadata built");
        assert_eq!(metadata.title.as_deref(), Some("Threat Report"));
        assert_eq!(metadata.subject.as_deref(), Some("Forensics"));
        assert_eq!(metadata.creator.as_deref(), Some("analyst"));
        assert_eq!(metadata.keywords.as_deref(), Some("malware,dde"));
        assert_eq!(metadata.description.as_deref(), Some("summary"));
        assert_eq!(metadata.last_modified_by.as_deref(), Some("reviewer"));
        assert_eq!(metadata.revision.as_deref(), Some("7"));
        assert_eq!(metadata.created.as_deref(), Some("2025-05-01T10:00:00Z"));
        assert_eq!(metadata.modified.as_deref(), Some("2025-05-01T11:00:00Z"));
        assert_eq!(metadata.category.as_deref(), Some("incident"));
        assert_eq!(metadata.content_status.as_deref(), Some("final"));
        assert_eq!(metadata.application.as_deref(), Some("docir"));
        assert_eq!(metadata.app_version.as_deref(), Some("1.0"));
        assert_eq!(metadata.company.as_deref(), Some("ACME"));
        assert_eq!(metadata.manager.as_deref(), Some("SecOps"));
        assert_eq!(metadata.custom_properties.len(), 6);
        assert!(matches!(
            metadata.custom_properties[0].value,
            PropertyValue::Integer(42)
        ));
        assert!(matches!(
            metadata.custom_properties[1].value,
            PropertyValue::Float(v) if (v - 3.25).abs() < f64::EPSILON
        ));
        assert!(matches!(
            metadata.custom_properties[2].value,
            PropertyValue::Boolean(true)
        ));
        assert!(matches!(
            metadata.custom_properties[3].value,
            PropertyValue::DateTime(ref v) if v == "2025-05-01T00:00:00Z"
        ));
        assert!(matches!(
            metadata.custom_properties[4].value,
            PropertyValue::Blob(ref v) if v == "QUJD"
        ));
        assert!(matches!(
            metadata.custom_properties[5].value,
            PropertyValue::String(ref v) if v == "alert"
        ));
    }

    #[test]
    fn parse_custom_properties_coerces_malformed_values_to_defaults() {
        let xml = r#"
            <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
                        xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
              <property pid="2" name="BadInt"><vt:int>not-a-number</vt:int></property>
              <property pid="3" name="BadFloat"><vt:r4>NaN?</vt:r4></property>
              <property pid="4" name="BoolNo"><vt:bool>false</vt:bool></property>
              <property pid="5" name="Unknown"><vt:custom>raw</vt:custom></property>
            </Properties>
        "#;
        let parser = OoxmlParser::new();
        let mut metadata = DocumentMetadata::new();
        parser.parse_custom_properties(xml, &mut metadata);

        assert_eq!(metadata.custom_properties.len(), 4);
        assert!(matches!(
            metadata.custom_properties[0].value,
            PropertyValue::Integer(0)
        ));
        assert!(matches!(
            metadata.custom_properties[1].value,
            PropertyValue::Float(v) if v == 0.0
        ));
        assert!(matches!(
            metadata.custom_properties[2].value,
            PropertyValue::Boolean(false)
        ));
        assert!(matches!(
            metadata.custom_properties[3].value,
            PropertyValue::String(ref v) if v == "raw"
        ));
    }
}
