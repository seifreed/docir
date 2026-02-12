use super::*;
use crate::ooxml::part_utils::read_xml_part_and_rels;
use crate::zip_handler::PackageReader;

impl OoxmlParser {
    pub(super) fn parse_pptx(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        content_types: &ContentTypes,
        metrics: &mut Option<ParseMetrics>,
    ) -> Result<ParsedDocument, ParseError> {
        let (presentation_xml, presentation_rels) = read_xml_part_and_rels(zip, main_part_path)?;

        let mut parser = PptxParser::new();
        let root_id = parser.parse_presentation(
            zip,
            &presentation_xml,
            &presentation_rels,
            main_part_path,
        )?;
        let mut store = parser.into_store();

        let metadata_id = self.parse_metadata(zip)?;
        if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
            doc.metadata = metadata_id;
        }
        if metadata_id.is_some() {
            if let Some(metadata) = self.build_metadata(zip) {
                store.insert(IRNode::Metadata(metadata));
            }
        }

        let start = std::time::Instant::now();
        self.parse_shared_parts(zip, content_types, &mut store, root_id)?;
        if let Some(m) = metrics.as_mut() {
            m.shared_parts_ms = start.elapsed().as_millis();
        }

        if self.config.scan_security_on_parse {
            let start = std::time::Instant::now();
            let scanner = security::SecurityScanner::new(&self.config);
            scanner.scan_zip(zip, &mut store)?;
            if let Some(m) = metrics.as_mut() {
                m.security_scan_ms = start.elapsed().as_millis();
            }
        }

        self.post_process_ooxml(zip, content_types, &mut store, root_id, metrics)?;

        Ok(ParsedDocument {
            root_id,
            format: DocumentFormat::Presentation,
            store,
            metrics: None,
        })
    }
}
