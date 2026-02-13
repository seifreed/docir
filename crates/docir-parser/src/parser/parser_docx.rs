use super::*;
use crate::ooxml::part_utils::read_xml_part_and_rels;
use crate::zip_handler::PackageReader;

impl OoxmlParser {
    /// Parse a DOCX document.
    pub(super) fn parse_docx(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        content_types: &ContentTypes,
        metrics: &mut Option<ParseMetrics>,
    ) -> Result<ParsedDocument, ParseError> {
        // Read main document + relationships
        let (document_xml, doc_rels) = read_xml_part_and_rels(zip, main_part_path)?;

        // Parse document
        let mut docx_parser = DocxParser::new();

        let header_footer_map =
            self.parse_docx_headers_footers(zip, main_part_path, &doc_rels, &mut docx_parser)?;

        let parts = self.parse_docx_word_parts(zip, main_part_path, &doc_rels, &mut docx_parser);

        let root_id =
            docx_parser.parse_document(&document_xml, &doc_rels, Some(&header_footer_map))?;
        let mut store = docx_parser.into_store();

        if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
            doc.styles = parts.styles_id;
            doc.styles_with_effects = parts.styles_with_effects_id;
            doc.numbering = parts.numbering_id;
            doc.comments = parts.comments;
            doc.footnotes = parts.footnotes;
            doc.endnotes = parts.endnotes;
            doc.settings = parts.settings_id;
            doc.web_settings = parts.web_settings_id;
            doc.font_table = parts.font_table_id;
            doc.comments_extended = parts.comments_ext_id;
            doc.comment_id_map = parts.comments_id_map_id;
            if let Some(glossary_id) = parts.glossary_id {
                doc.shared_parts.push(glossary_id);
            }
        }

        self.finalize_ooxml_document(zip, content_types, &mut store, root_id, metrics)?;

        Ok(self.build_parsed_document(root_id, DocumentFormat::WordProcessing, store))
    }
}
