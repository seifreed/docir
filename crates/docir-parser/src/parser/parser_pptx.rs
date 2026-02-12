use super::*;
use crate::ooxml::part_utils::read_relationships;

impl OoxmlParser {
    pub(super) fn parse_pptx<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        main_part_path: &str,
        content_types: &ContentTypes,
        metrics: &mut Option<ParseMetrics>,
    ) -> Result<ParsedDocument, ParseError> {
        let presentation_xml = zip.read_file_string(main_part_path)?;

        let presentation_rels = read_relationships(zip, main_part_path)?;

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

        let start = std::time::Instant::now();
        self.scan_security_content(zip, &mut store, root_id, content_types)?;
        if let Some(m) = metrics.as_mut() {
            m.security_scan_ms = start.elapsed().as_millis();
        }

        // Link shapes/animations to shared parts (charts, SmartArt, media, OLE)
        self.link_shapes_to_shared_parts(&mut store);

        let start = std::time::Instant::now();
        self.add_extension_parts_and_diagnostics(zip, content_types, &mut store, root_id)?;
        if let Some(m) = metrics.as_mut() {
            m.extension_parts_ms = start.elapsed().as_millis();
        }

        let start = std::time::Instant::now();
        normalize_store(&mut store, root_id);
        if let Some(m) = metrics.as_mut() {
            m.normalization_ms = start.elapsed().as_millis();
        }

        Ok(ParsedDocument {
            root_id,
            format: DocumentFormat::Presentation,
            store,
            metrics: None,
        })
    }
}
