use crate::ir::*;
use crate::security::*;
use std::collections::HashMap;

use super::{IrVisitor, VisitControl, VisitorResult};

/// A simple visitor that collects all text content.
pub struct TextCollector {
    pub text: String,
}

impl TextCollector {
    pub fn new() -> Self {
        Self {
            text: String::new(),
        }
    }
}

impl Default for TextCollector {
    fn default() -> Self {
        Self::new()
    }
}

impl IrVisitor for TextCollector {
    fn visit_run(&mut self, run: &Run) -> VisitorResult<VisitControl> {
        self.text.push_str(&run.text);
        Ok(VisitControl::Continue)
    }

    fn visit_paragraph(&mut self, _para: &Paragraph) -> VisitorResult<VisitControl> {
        if !self.text.is_empty() && !self.text.ends_with('\n') {
            self.text.push('\n');
        }
        Ok(VisitControl::Continue)
    }
}

/// A visitor that counts nodes by type.
pub struct NodeCounter {
    pub counts: HashMap<String, usize>,
}

impl NodeCounter {
    pub fn new() -> Self {
        Self {
            counts: HashMap::new(),
        }
    }

    fn increment(&mut self, node_type: &str) {
        *self.counts.entry(node_type.to_string()).or_insert(0) += 1;
    }
}

impl Default for NodeCounter {
    fn default() -> Self {
        Self::new()
    }
}

impl IrVisitor for NodeCounter {
    fn visit_document(&mut self, _: &Document) -> VisitorResult<VisitControl> {
        self.increment("Document");
        Ok(VisitControl::Continue)
    }

    fn visit_section(&mut self, _: &Section) -> VisitorResult<VisitControl> {
        self.increment("Section");
        Ok(VisitControl::Continue)
    }

    fn visit_paragraph(&mut self, _: &Paragraph) -> VisitorResult<VisitControl> {
        self.increment("Paragraph");
        Ok(VisitControl::Continue)
    }

    fn visit_run(&mut self, _: &Run) -> VisitorResult<VisitControl> {
        self.increment("Run");
        Ok(VisitControl::Continue)
    }

    fn visit_table(&mut self, _: &Table) -> VisitorResult<VisitControl> {
        self.increment("Table");
        Ok(VisitControl::Continue)
    }

    fn visit_table_row(&mut self, _: &TableRow) -> VisitorResult<VisitControl> {
        self.increment("TableRow");
        Ok(VisitControl::Continue)
    }

    fn visit_table_cell(&mut self, _: &TableCell) -> VisitorResult<VisitControl> {
        self.increment("TableCell");
        Ok(VisitControl::Continue)
    }

    fn visit_slide(&mut self, _: &Slide) -> VisitorResult<VisitControl> {
        self.increment("Slide");
        Ok(VisitControl::Continue)
    }

    fn visit_worksheet(&mut self, _: &Worksheet) -> VisitorResult<VisitControl> {
        self.increment("Worksheet");
        Ok(VisitControl::Continue)
    }

    fn visit_cell(&mut self, _: &Cell) -> VisitorResult<VisitControl> {
        self.increment("Cell");
        Ok(VisitControl::Continue)
    }

    fn visit_calc_chain(&mut self, _: &CalcChain) -> VisitorResult<VisitControl> {
        self.increment("CalcChain");
        Ok(VisitControl::Continue)
    }

    fn visit_sheet_comment(&mut self, _: &SheetComment) -> VisitorResult<VisitControl> {
        self.increment("SheetComment");
        Ok(VisitControl::Continue)
    }

    fn visit_sheet_metadata(&mut self, _: &SheetMetadata) -> VisitorResult<VisitControl> {
        self.increment("SheetMetadata");
        Ok(VisitControl::Continue)
    }

    fn visit_macro_project(&mut self, _: &MacroProject) -> VisitorResult<VisitControl> {
        self.increment("MacroProject");
        Ok(VisitControl::Continue)
    }

    fn visit_ole_object(&mut self, _: &OleObject) -> VisitorResult<VisitControl> {
        self.increment("OleObject");
        Ok(VisitControl::Continue)
    }

    fn visit_external_ref(&mut self, _: &ExternalReference) -> VisitorResult<VisitControl> {
        self.increment("ExternalReference");
        Ok(VisitControl::Continue)
    }

    fn visit_style_set(&mut self, _: &StyleSet) -> VisitorResult<VisitControl> {
        self.increment("StyleSet");
        Ok(VisitControl::Continue)
    }

    fn visit_numbering_set(&mut self, _: &NumberingSet) -> VisitorResult<VisitControl> {
        self.increment("NumberingSet");
        Ok(VisitControl::Continue)
    }

    fn visit_comment(&mut self, _: &Comment) -> VisitorResult<VisitControl> {
        self.increment("Comment");
        Ok(VisitControl::Continue)
    }

    fn visit_footnote(&mut self, _: &Footnote) -> VisitorResult<VisitControl> {
        self.increment("Footnote");
        Ok(VisitControl::Continue)
    }

    fn visit_endnote(&mut self, _: &Endnote) -> VisitorResult<VisitControl> {
        self.increment("Endnote");
        Ok(VisitControl::Continue)
    }

    fn visit_header(&mut self, _: &Header) -> VisitorResult<VisitControl> {
        self.increment("Header");
        Ok(VisitControl::Continue)
    }

    fn visit_footer(&mut self, _: &Footer) -> VisitorResult<VisitControl> {
        self.increment("Footer");
        Ok(VisitControl::Continue)
    }

    fn visit_word_settings(&mut self, _: &WordSettings) -> VisitorResult<VisitControl> {
        self.increment("WordSettings");
        Ok(VisitControl::Continue)
    }

    fn visit_web_settings(&mut self, _: &WebSettings) -> VisitorResult<VisitControl> {
        self.increment("WebSettings");
        Ok(VisitControl::Continue)
    }

    fn visit_font_table(&mut self, _: &FontTable) -> VisitorResult<VisitControl> {
        self.increment("FontTable");
        Ok(VisitControl::Continue)
    }

    fn visit_content_control(&mut self, _: &ContentControl) -> VisitorResult<VisitControl> {
        self.increment("ContentControl");
        Ok(VisitControl::Continue)
    }

    fn visit_people_part(&mut self, _: &PeoplePart) -> VisitorResult<VisitControl> {
        self.increment("PeoplePart");
        Ok(VisitControl::Continue)
    }

    fn visit_web_extension(&mut self, _: &WebExtension) -> VisitorResult<VisitControl> {
        self.increment("WebExtension");
        Ok(VisitControl::Continue)
    }

    fn visit_web_extension_taskpane(
        &mut self,
        _: &WebExtensionTaskpane,
    ) -> VisitorResult<VisitControl> {
        self.increment("WebExtensionTaskpane");
        Ok(VisitControl::Continue)
    }

    fn visit_glossary_document(&mut self, _: &GlossaryDocument) -> VisitorResult<VisitControl> {
        self.increment("GlossaryDocument");
        Ok(VisitControl::Continue)
    }

    fn visit_glossary_entry(&mut self, _: &GlossaryEntry) -> VisitorResult<VisitControl> {
        self.increment("GlossaryEntry");
        Ok(VisitControl::Continue)
    }

    fn visit_vml_drawing(&mut self, _: &VmlDrawing) -> VisitorResult<VisitControl> {
        self.increment("VmlDrawing");
        Ok(VisitControl::Continue)
    }

    fn visit_vml_shape(&mut self, _: &VmlShape) -> VisitorResult<VisitControl> {
        self.increment("VmlShape");
        Ok(VisitControl::Continue)
    }

    fn visit_drawing_part(&mut self, _: &DrawingPart) -> VisitorResult<VisitControl> {
        self.increment("DrawingPart");
        Ok(VisitControl::Continue)
    }

    fn visit_external_link_part(&mut self, _: &ExternalLinkPart) -> VisitorResult<VisitControl> {
        self.increment("ExternalLinkPart");
        Ok(VisitControl::Continue)
    }

    fn visit_connection_part(&mut self, _: &ConnectionPart) -> VisitorResult<VisitControl> {
        self.increment("ConnectionPart");
        Ok(VisitControl::Continue)
    }

    fn visit_slicer_part(&mut self, _: &SlicerPart) -> VisitorResult<VisitControl> {
        self.increment("SlicerPart");
        Ok(VisitControl::Continue)
    }

    fn visit_timeline_part(&mut self, _: &TimelinePart) -> VisitorResult<VisitControl> {
        self.increment("TimelinePart");
        Ok(VisitControl::Continue)
    }

    fn visit_query_table_part(&mut self, _: &QueryTablePart) -> VisitorResult<VisitControl> {
        self.increment("QueryTablePart");
        Ok(VisitControl::Continue)
    }

    fn visit_presentation_info(&mut self, _: &PresentationInfo) -> VisitorResult<VisitControl> {
        self.increment("PresentationInfo");
        Ok(VisitControl::Continue)
    }

    fn visit_bookmark_start(&mut self, _: &BookmarkStart) -> VisitorResult<VisitControl> {
        self.increment("BookmarkStart");
        Ok(VisitControl::Continue)
    }

    fn visit_bookmark_end(&mut self, _: &BookmarkEnd) -> VisitorResult<VisitControl> {
        self.increment("BookmarkEnd");
        Ok(VisitControl::Continue)
    }

    fn visit_field(&mut self, _: &Field) -> VisitorResult<VisitControl> {
        self.increment("Field");
        Ok(VisitControl::Continue)
    }

    fn visit_revision(&mut self, _: &Revision) -> VisitorResult<VisitControl> {
        self.increment("Revision");
        Ok(VisitControl::Continue)
    }

    fn visit_comment_extension_set(
        &mut self,
        _: &CommentExtensionSet,
    ) -> VisitorResult<VisitControl> {
        self.increment("CommentExtensionSet");
        Ok(VisitControl::Continue)
    }

    fn visit_comment_id_map(&mut self, _: &CommentIdMap) -> VisitorResult<VisitControl> {
        self.increment("CommentIdMap");
        Ok(VisitControl::Continue)
    }

    fn visit_slide_master(&mut self, _: &SlideMaster) -> VisitorResult<VisitControl> {
        self.increment("SlideMaster");
        Ok(VisitControl::Continue)
    }

    fn visit_slide_layout(&mut self, _: &SlideLayout) -> VisitorResult<VisitControl> {
        self.increment("SlideLayout");
        Ok(VisitControl::Continue)
    }

    fn visit_notes_master(&mut self, _: &NotesMaster) -> VisitorResult<VisitControl> {
        self.increment("NotesMaster");
        Ok(VisitControl::Continue)
    }

    fn visit_handout_master(&mut self, _: &HandoutMaster) -> VisitorResult<VisitControl> {
        self.increment("HandoutMaster");
        Ok(VisitControl::Continue)
    }

    fn visit_notes_slide(&mut self, _: &NotesSlide) -> VisitorResult<VisitControl> {
        self.increment("NotesSlide");
        Ok(VisitControl::Continue)
    }

    fn visit_worksheet_drawing(&mut self, _: &WorksheetDrawing) -> VisitorResult<VisitControl> {
        self.increment("WorksheetDrawing");
        Ok(VisitControl::Continue)
    }

    fn visit_chart_data(&mut self, _: &ChartData) -> VisitorResult<VisitControl> {
        self.increment("ChartData");
        Ok(VisitControl::Continue)
    }
    fn visit_metadata(&mut self, _: &DocumentMetadata) -> VisitorResult<VisitControl> {
        self.increment("Metadata");
        Ok(VisitControl::Continue)
    }

    fn visit_theme(&mut self, _: &Theme) -> VisitorResult<VisitControl> {
        self.increment("Theme");
        Ok(VisitControl::Continue)
    }

    fn visit_media_asset(&mut self, _: &MediaAsset) -> VisitorResult<VisitControl> {
        self.increment("MediaAsset");
        Ok(VisitControl::Continue)
    }

    fn visit_custom_xml_part(&mut self, _: &CustomXmlPart) -> VisitorResult<VisitControl> {
        self.increment("CustomXmlPart");
        Ok(VisitControl::Continue)
    }

    fn visit_relationship_graph(&mut self, _: &RelationshipGraph) -> VisitorResult<VisitControl> {
        self.increment("RelationshipGraph");
        Ok(VisitControl::Continue)
    }

    fn visit_digital_signature(
        &mut self,
        _: &crate::ir::DigitalSignature,
    ) -> VisitorResult<VisitControl> {
        self.increment("DigitalSignature");
        Ok(VisitControl::Continue)
    }

    fn visit_extension_part(&mut self, _: &ExtensionPart) -> VisitorResult<VisitControl> {
        self.increment("ExtensionPart");
        Ok(VisitControl::Continue)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::security::ExternalRefType;

    #[test]
    fn text_collector_adds_newline_between_paragraphs() {
        let mut collector = TextCollector::new();
        collector.visit_run(&Run::new("hello")).unwrap();
        collector.visit_paragraph(&Paragraph::new()).unwrap();
        collector.visit_run(&Run::new("world")).unwrap();
        assert_eq!(collector.text, "hello\nworld");
    }

    #[test]
    fn node_counter_counts_supported_nodes() {
        let mut counter = NodeCounter::new();

        counter
            .visit_document(&Document::new(crate::types::DocumentFormat::WordProcessing))
            .unwrap();
        counter.visit_section(&Section::new()).unwrap();
        counter.visit_paragraph(&Paragraph::new()).unwrap();
        counter.visit_run(&Run::new("t")).unwrap();
        counter.visit_table(&Table::new()).unwrap();
        counter.visit_table_row(&TableRow::new()).unwrap();
        counter.visit_table_cell(&TableCell::new()).unwrap();
        counter.visit_slide(&Slide::new(1)).unwrap();
        counter.visit_worksheet(&Worksheet::new("Sheet1", 1)).unwrap();
        counter.visit_cell(&Cell::new("A1", 0, 0)).unwrap();
        counter.visit_calc_chain(&CalcChain::new()).unwrap();
        counter
            .visit_sheet_comment(&SheetComment::new("A1", "comment"))
            .unwrap();
        counter.visit_sheet_metadata(&SheetMetadata::new()).unwrap();
        counter.visit_macro_project(&MacroProject::new()).unwrap();
        counter.visit_ole_object(&OleObject::new()).unwrap();
        counter
            .visit_external_ref(&ExternalReference::new(
                ExternalRefType::Hyperlink,
                "https://example.test",
            ))
            .unwrap();
        counter.visit_style_set(&StyleSet::new()).unwrap();
        counter.visit_numbering_set(&NumberingSet::new()).unwrap();
        counter.visit_comment(&Comment::new("1")).unwrap();
        counter.visit_footnote(&Footnote::new("1")).unwrap();
        counter.visit_endnote(&Endnote::new("1")).unwrap();
        counter.visit_header(&Header::new()).unwrap();
        counter.visit_footer(&Footer::new()).unwrap();
        counter.visit_word_settings(&WordSettings::new()).unwrap();
        counter.visit_web_settings(&WebSettings::new()).unwrap();
        counter.visit_font_table(&FontTable::new()).unwrap();
        counter.visit_content_control(&ContentControl::new()).unwrap();
        counter.visit_people_part(&PeoplePart::new()).unwrap();
        counter.visit_web_extension(&WebExtension::new()).unwrap();
        counter
            .visit_web_extension_taskpane(&WebExtensionTaskpane::new())
            .unwrap();
        counter
            .visit_glossary_document(&GlossaryDocument::new())
            .unwrap();
        counter.visit_glossary_entry(&GlossaryEntry::new()).unwrap();
        counter.visit_vml_drawing(&VmlDrawing::new("vml.xml")).unwrap();
        counter.visit_vml_shape(&VmlShape::new()).unwrap();
        counter
            .visit_drawing_part(&DrawingPart::new("drawing.xml"))
            .unwrap();
        counter
            .visit_external_link_part(&ExternalLinkPart::new())
            .unwrap();
        counter
            .visit_connection_part(&ConnectionPart::new())
            .unwrap();
        counter.visit_slicer_part(&SlicerPart::new()).unwrap();
        counter.visit_timeline_part(&TimelinePart::new()).unwrap();
        counter
            .visit_query_table_part(&QueryTablePart::new())
            .unwrap();
        counter
            .visit_presentation_info(&PresentationInfo::new())
            .unwrap();
        counter.visit_bookmark_start(&BookmarkStart::new("b1")).unwrap();
        counter.visit_bookmark_end(&BookmarkEnd::new("b1")).unwrap();
        counter.visit_field(&Field::new(Some("DATE".to_string()))).unwrap();
        counter
            .visit_revision(&Revision::new(RevisionType::Insert))
            .unwrap();
        counter
            .visit_comment_extension_set(&CommentExtensionSet::new())
            .unwrap();
        counter.visit_comment_id_map(&CommentIdMap::new()).unwrap();
        counter.visit_slide_master(&SlideMaster::new()).unwrap();
        counter.visit_slide_layout(&SlideLayout::new()).unwrap();
        counter.visit_notes_master(&NotesMaster::new()).unwrap();
        counter.visit_handout_master(&HandoutMaster::new()).unwrap();
        counter.visit_notes_slide(&NotesSlide::new()).unwrap();
        counter
            .visit_worksheet_drawing(&WorksheetDrawing::new())
            .unwrap();
        counter.visit_chart_data(&ChartData::new()).unwrap();
        counter.visit_metadata(&DocumentMetadata::new()).unwrap();
        counter.visit_theme(&Theme::new()).unwrap();
        counter
            .visit_media_asset(&MediaAsset::new(
                "m.bin",
                crate::ir::MediaType::Other,
                1,
            ))
            .unwrap();
        counter
            .visit_custom_xml_part(&CustomXmlPart::new("custom.xml", 1))
            .unwrap();
        counter
            .visit_relationship_graph(&RelationshipGraph::new("xl/workbook.xml"))
            .unwrap();
        counter
            .visit_digital_signature(&crate::ir::DigitalSignature::new())
            .unwrap();
        counter
            .visit_extension_part(&ExtensionPart::new(
                "ext.bin",
                7,
                crate::ir::ExtensionPartKind::Unknown,
            ))
            .unwrap();

        assert_eq!(counter.counts.get("Document"), Some(&1));
        assert_eq!(counter.counts.get("Run"), Some(&1));
        assert_eq!(counter.counts.get("ExtensionPart"), Some(&1));
    }
}
