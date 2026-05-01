use crate::ir::*;
#[cfg(test)]
use crate::security::*;
use std::collections::HashMap;

use super::{IrVisitor, VisitControl, VisitorResult};
#[path = "visitors_node_counter.rs"]
mod visitors_node_counter;

/// A simple visitor that collects all text content.
pub struct TextCollector {
    pub text: String,
}

impl TextCollector {
    /// Public API entrypoint: new.
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
    /// Public API entrypoint: new.
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

    fn assert_counts(counter: &NodeCounter, expected: &[(&str, usize)]) {
        for (key, value) in expected {
            assert_eq!(counter.counts.get(*key), Some(value));
        }
    }

    fn visit_textual_nodes(counter: &mut NodeCounter) {
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
        counter
            .visit_worksheet(&Worksheet::new("Sheet1", 1))
            .unwrap();
        counter.visit_cell(&Cell::new("A1", 0, 0)).unwrap();
        counter.visit_calc_chain(&CalcChain::new()).unwrap();
        counter
            .visit_sheet_comment(&SheetComment::new("A1", "comment"))
            .unwrap();
        counter.visit_sheet_metadata(&SheetMetadata::new()).unwrap();
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
        counter
            .visit_content_control(&ContentControl::new())
            .unwrap();
        counter.visit_people_part(&PeoplePart::new()).unwrap();
        counter
            .visit_bookmark_start(&BookmarkStart::new("b1"))
            .unwrap();
        counter.visit_bookmark_end(&BookmarkEnd::new("b1")).unwrap();
        counter
            .visit_field(&Field::new(Some("DATE".to_string())))
            .unwrap();
        counter
            .visit_revision(&Revision::new(RevisionType::Insert))
            .unwrap();
    }

    fn visit_auxiliary_nodes(counter: &mut NodeCounter) {
        counter
            .visit_comment_extension_set(&CommentExtensionSet::new())
            .unwrap();
        counter.visit_comment_id_map(&CommentIdMap::new()).unwrap();
        counter.visit_macro_project(&MacroProject::new()).unwrap();
        counter.visit_ole_object(&OleObject::new()).unwrap();
        counter
            .visit_external_ref(&ExternalReference::new(
                ExternalRefType::Hyperlink,
                "https://example.test",
            ))
            .unwrap();
        counter.visit_web_extension(&WebExtension::new()).unwrap();
        counter
            .visit_web_extension_taskpane(&WebExtensionTaskpane::new())
            .unwrap();
        counter
            .visit_glossary_document(&GlossaryDocument::new())
            .unwrap();
        counter.visit_glossary_entry(&GlossaryEntry::new()).unwrap();
        counter
            .visit_vml_drawing(&VmlDrawing::new("vml.xml"))
            .unwrap();
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
            .visit_media_asset(&MediaAsset::new("m.bin", crate::ir::MediaType::Other, 1))
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
    }

    #[test]
    fn node_counter_counts_textual_nodes() {
        let mut counter = NodeCounter::new();
        visit_textual_nodes(&mut counter);

        assert_counts(
            &counter,
            &[
                ("Document", 1),
                ("Section", 1),
                ("Paragraph", 1),
                ("Run", 1),
                ("Table", 1),
                ("TableRow", 1),
                ("TableCell", 1),
                ("Slide", 1),
                ("Worksheet", 1),
                ("Cell", 1),
                ("CalcChain", 1),
                ("SheetComment", 1),
                ("SheetMetadata", 1),
                ("StyleSet", 1),
                ("NumberingSet", 1),
                ("Comment", 1),
                ("Footnote", 1),
                ("Endnote", 1),
                ("Header", 1),
                ("Footer", 1),
                ("WordSettings", 1),
                ("WebSettings", 1),
                ("FontTable", 1),
                ("ContentControl", 1),
                ("PeoplePart", 1),
                ("BookmarkStart", 1),
                ("BookmarkEnd", 1),
                ("Field", 1),
                ("Revision", 1),
            ],
        );
    }

    #[test]
    fn node_counter_counts_auxiliary_nodes() {
        let mut counter = NodeCounter::new();
        visit_auxiliary_nodes(&mut counter);

        assert_counts(
            &counter,
            &[
                ("CommentExtensionSet", 1),
                ("CommentIdMap", 1),
                ("MacroProject", 1),
                ("OleObject", 1),
                ("ExternalReference", 1),
                ("WebExtension", 1),
                ("WebExtensionTaskpane", 1),
                ("GlossaryDocument", 1),
                ("GlossaryEntry", 1),
                ("VmlDrawing", 1),
                ("VmlShape", 1),
                ("DrawingPart", 1),
                ("ExternalLinkPart", 1),
                ("ConnectionPart", 1),
                ("SlicerPart", 1),
                ("TimelinePart", 1),
                ("QueryTablePart", 1),
                ("PresentationInfo", 1),
                ("SlideMaster", 1),
                ("SlideLayout", 1),
                ("NotesMaster", 1),
                ("HandoutMaster", 1),
                ("NotesSlide", 1),
                ("WorksheetDrawing", 1),
                ("ChartData", 1),
                ("Metadata", 1),
                ("Theme", 1),
                ("MediaAsset", 1),
                ("CustomXmlPart", 1),
                ("RelationshipGraph", 1),
                ("DigitalSignature", 1),
                ("ExtensionPart", 1),
            ],
        );
    }

    #[test]
    fn defaults_and_empty_paragraph_do_not_modify_text() {
        let mut collector = TextCollector::default();
        collector.visit_paragraph(&Paragraph::new()).unwrap();
        assert!(collector.text.is_empty());

        let mut with_newline = TextCollector {
            text: "line\n".to_string(),
        };
        with_newline.visit_paragraph(&Paragraph::new()).unwrap();
        assert_eq!(with_newline.text, "line\n");

        let counter = NodeCounter::default();
        assert!(counter.counts.is_empty());
    }
}
