use crate::ooxml::relationships::Relationships;
use crate::ooxml::xlsx::{ParseError, PivotCache, XlsxParser};
use crate::zip_handler::PackageReader;
use docir_core::types::NodeId;

#[path = "worksheet_parse_external_chartsheet.rs"]
mod worksheet_parse_external_chartsheet;
#[path = "worksheet_parse_external_links.rs"]
mod worksheet_parse_external_links;
#[path = "worksheet_parse_external_pivot.rs"]
mod worksheet_parse_external_pivot;

impl XlsxParser {
    pub(super) fn parse_chartsheet(
        &mut self,
        zip: &mut impl PackageReader,
        xml: &str,
        sheet_path: &str,
        relationships: &Relationships,
    ) -> Result<Option<NodeId>, ParseError> {
        worksheet_parse_external_chartsheet::parse_chartsheet_impl(
            self,
            zip,
            xml,
            sheet_path,
            relationships,
        )
    }

    pub(crate) fn parse_external_links_and_connections(
        &mut self,
        zip: &mut impl PackageReader,
        workbook_path: &str,
        workbook_rels: &Relationships,
    ) -> Result<(), ParseError> {
        worksheet_parse_external_links::parse_external_links_and_connections_impl(
            self,
            zip,
            workbook_path,
            workbook_rels,
        )
    }

    pub(crate) fn parse_pivot_cache(
        &mut self,
        zip: &mut impl PackageReader,
        xml: &str,
        cache_path: &str,
        cache_id: u32,
    ) -> Result<PivotCache, ParseError> {
        worksheet_parse_external_pivot::parse_pivot_cache_impl(self, zip, xml, cache_path, cache_id)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::zip_handler::PackageReader;
    use std::collections::HashMap;

    #[derive(Default)]
    struct MockPackageReader {
        files: HashMap<String, Vec<u8>>,
    }

    impl MockPackageReader {
        fn insert_str(&mut self, path: &str, data: &str) {
            self.files
                .insert(path.to_string(), data.as_bytes().to_vec());
        }
    }

    impl PackageReader for MockPackageReader {
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
            Ok(self.read_file(name)?.len() as u64)
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
    fn parse_chartsheet_returns_none_when_no_chart_rel_is_present() {
        let mut parser = XlsxParser::new();
        let mut zip = MockPackageReader::default();
        let xml = r#"<chartsheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews/></chartsheet>"#;
        let rels = Relationships::default();
        let drawing = parser
            .parse_chartsheet(&mut zip, xml, "xl/chartsheets/sheet1.xml", &rels)
            .expect("chartsheet parse");
        assert!(drawing.is_none());
    }

    #[test]
    fn parse_chartsheet_builds_drawing_when_relationship_and_chart_exist() {
        let mut parser = XlsxParser::new();
        let mut zip = MockPackageReader::default();
        zip.insert_str(
            "xl/charts/chart1.xml",
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#,
        );

        let mut rels = Relationships::default();
        rels.by_id.insert(
            "rIdChart1".to_string(),
            crate::ooxml::relationships::Relationship {
                id: "rIdChart1".to_string(),
                rel_type: crate::ooxml::xlsx::rel_type::CHART.to_string(),
                target: "../charts/chart1.xml".to_string(),
                target_mode: crate::ooxml::relationships::TargetMode::Internal,
            },
        );

        let xml = r#"
            <chartsheet xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <drawing>
                <c:chart r:id="rIdChart1"/>
              </drawing>
            </chartsheet>
        "#;

        let drawing_id = parser
            .parse_chartsheet(&mut zip, xml, "xl/chartsheets/sheet1.xml", &rels)
            .expect("chartsheet parse")
            .expect("drawing should be created");

        let drawing = match parser.store.get(drawing_id) {
            Some(crate::ooxml::xlsx::IRNode::WorksheetDrawing(drawing)) => drawing,
            other => panic!("expected worksheet drawing, got {other:?}"),
        };
        assert_eq!(drawing.shapes.len(), 1);

        let shape = match parser.store.get(drawing.shapes[0]) {
            Some(crate::ooxml::xlsx::IRNode::Shape(shape)) => shape,
            other => panic!("expected shape, got {other:?}"),
        };
        assert_eq!(shape.media_target.as_deref(), Some("xl/charts/chart1.xml"));
        assert_eq!(shape.relationship_id.as_deref(), Some("rIdChart1"));
    }

    #[test]
    fn parse_pivot_cache_reads_cache_source_without_records_part() {
        let mut parser = XlsxParser::new();
        let mut zip = MockPackageReader::default();
        let cache_xml = r#"
            <pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <cacheSource type="worksheet"/>
              <worksheetSource sheet="Sheet1" ref="A1:B3"/>
            </pivotCacheDefinition>
        "#;
        zip.insert_str("xl/pivotCache/pivotCacheDefinition1.xml", cache_xml);
        let cache = parser
            .parse_pivot_cache(
                &mut zip,
                cache_xml,
                "xl/pivotCache/pivotCacheDefinition1.xml",
                7,
            )
            .expect("pivot cache parse");
        assert_eq!(
            cache.cache_source.as_deref(),
            Some("worksheet:Sheet1!A1:B3")
        );
        assert_eq!(cache.cache_id, 7);
        assert!(cache.records.is_none());
    }

    #[test]
    fn parse_pivot_cache_reads_connection_cache_source() {
        let mut parser = XlsxParser::new();
        let mut zip = MockPackageReader::default();
        let cache_xml = r#"
            <pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <cacheSource type="external" connectionId="42"/>
            </pivotCacheDefinition>
        "#;
        zip.insert_str("xl/pivotCache/pivotCacheDefinition2.xml", cache_xml);

        let cache = parser
            .parse_pivot_cache(
                &mut zip,
                cache_xml,
                "xl/pivotCache/pivotCacheDefinition2.xml",
                8,
            )
            .expect("pivot cache parse");

        assert_eq!(cache.cache_source.as_deref(), Some("connection:42"));
        assert_eq!(cache.cache_id, 8);
    }

    #[test]
    fn parse_external_links_and_connections_collects_shared_parts_and_external_refs() {
        let mut parser = XlsxParser::new();
        let mut zip = MockPackageReader::default();

        zip.insert_str(
            "xl/externalLinks/externalLink1.xml",
            r#"<externalLink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rIdBook"/></externalLink>"#,
        );
        zip.insert_str(
            "xl/externalLinks/_rels/externalLink1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdBook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="https://example.test/book.xlsx" TargetMode="External"/></Relationships>"#,
        );
        zip.insert_str(
            "xl/connections.xml",
            r#"<connections><connection id="1" name="Conn1"><dbPr connection="Provider=SQLOLEDB;Data Source=example.test;"/></connection></connections>"#,
        );

        let mut workbook_rels = Relationships::default();
        workbook_rels.by_id.insert(
            "rIdExt1".to_string(),
            crate::ooxml::relationships::Relationship {
                id: "rIdExt1".to_string(),
                rel_type: crate::ooxml::xlsx::rel_type::EXTERNAL_LINK.to_string(),
                target: "externalLinks/externalLink1.xml".to_string(),
                target_mode: crate::ooxml::relationships::TargetMode::Internal,
            },
        );
        workbook_rels.by_type.insert(
            crate::ooxml::xlsx::rel_type::EXTERNAL_LINK.to_string(),
            vec!["rIdExt1".to_string()],
        );

        parser
            .parse_external_links_and_connections(&mut zip, "xl/workbook.xml", &workbook_rels)
            .expect("parse external links and connections");

        let mut external_link_parts = 0usize;
        let mut connection_parts = 0usize;
        for node in parser.store.values() {
            match node {
                crate::ooxml::xlsx::IRNode::ExternalLinkPart(_) => external_link_parts += 1,
                crate::ooxml::xlsx::IRNode::ConnectionPart(_) => connection_parts += 1,
                _ => {}
            }
        }
        assert_eq!(external_link_parts, 1);
        assert_eq!(connection_parts, 1);
        assert!(parser.security_info.external_refs.len() >= 2);
    }

    #[test]
    fn parse_chartsheet_returns_none_when_chart_file_is_missing() {
        let mut parser = XlsxParser::new();
        let mut zip = MockPackageReader::default();
        let mut rels = Relationships::default();
        rels.by_id.insert(
            "rIdChartMissing".to_string(),
            crate::ooxml::relationships::Relationship {
                id: "rIdChartMissing".to_string(),
                rel_type: crate::ooxml::xlsx::rel_type::CHART.to_string(),
                target: "../charts/missing.xml".to_string(),
                target_mode: crate::ooxml::relationships::TargetMode::Internal,
            },
        );

        let xml = r#"
            <chartsheet xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <drawing>
                <c:chart r:id="rIdChartMissing"/>
              </drawing>
            </chartsheet>
        "#;

        let drawing = parser
            .parse_chartsheet(&mut zip, xml, "xl/chartsheets/sheet2.xml", &rels)
            .expect("chartsheet parse");
        assert!(drawing.is_none());
    }
}
