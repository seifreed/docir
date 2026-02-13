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

        self.finalize_ooxml_document(zip, content_types, &mut store, root_id, metrics)?;

        Ok(self.build_parsed_document(root_id, DocumentFormat::Presentation, store))
    }
}
