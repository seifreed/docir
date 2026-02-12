use crate::ir::*;
use crate::types::NodeId;

use super::{IrStore, IrVisitor, VisitControl, VisitorResult};

/// Walks the IR tree in pre-order (parent before children).
pub struct PreOrderWalker<'a> {
    store: &'a IrStore,
    stack: Vec<NodeId>,
}

impl<'a> PreOrderWalker<'a> {
    /// Creates a new PreOrderWalker starting from the given root node.
    pub fn new(store: &'a IrStore, root: NodeId) -> Self {
        Self {
            store,
            stack: vec![root],
        }
    }

    /// Walks the tree, calling the visitor for each node.
    pub fn walk<V: IrVisitor>(&mut self, visitor: &mut V) -> VisitorResult<()> {
        while let Some(node_id) = self.stack.pop() {
            let Some(node) = self.store.get(node_id) else {
                continue;
            };

            let control = self.visit_node(visitor, node)?;

            match control {
                VisitControl::Stop => return Ok(()),
                VisitControl::SkipChildren => continue,
                VisitControl::Continue => {
                    // Push children in reverse order so they're visited in order
                    let children = self.get_children(node);
                    for child_id in children.into_iter().rev() {
                        self.stack.push(child_id);
                    }
                }
            }
        }
        Ok(())
    }

    fn visit_node<V: IrVisitor>(
        &self,
        visitor: &mut V,
        node: &IRNode,
    ) -> VisitorResult<VisitControl> {
        match node {
            IRNode::Document(n) => visitor.visit_document(n),
            IRNode::Section(n) => visitor.visit_section(n),
            IRNode::Paragraph(n) => visitor.visit_paragraph(n),
            IRNode::Run(n) => visitor.visit_run(n),
            IRNode::Hyperlink(n) => visitor.visit_hyperlink(n),
            IRNode::Table(n) => visitor.visit_table(n),
            IRNode::TableRow(n) => visitor.visit_table_row(n),
            IRNode::TableCell(n) => visitor.visit_table_cell(n),
            IRNode::Slide(n) => visitor.visit_slide(n),
            IRNode::Shape(n) => visitor.visit_shape(n),
            IRNode::Worksheet(n) => visitor.visit_worksheet(n),
            IRNode::Cell(n) => visitor.visit_cell(n),
            IRNode::SharedStringTable(n) => visitor.visit_shared_string_table(n),
            IRNode::SpreadsheetStyles(n) => visitor.visit_spreadsheet_styles(n),
            IRNode::DefinedName(n) => visitor.visit_defined_name(n),
            IRNode::ConditionalFormat(n) => visitor.visit_conditional_format(n),
            IRNode::DataValidation(n) => visitor.visit_data_validation(n),
            IRNode::TableDefinition(n) => visitor.visit_table_definition(n),
            IRNode::PivotTable(n) => visitor.visit_pivot_table(n),
            IRNode::PivotCache(n) => visitor.visit_pivot_cache(n),
            IRNode::PivotCacheRecords(n) => visitor.visit_pivot_cache_records(n),
            IRNode::CalcChain(n) => visitor.visit_calc_chain(n),
            IRNode::SheetComment(n) => visitor.visit_sheet_comment(n),
            IRNode::SheetMetadata(n) => visitor.visit_sheet_metadata(n),
            IRNode::WorkbookProperties(n) => visitor.visit_workbook_properties(n),
            IRNode::MacroProject(n) => visitor.visit_macro_project(n),
            IRNode::MacroModule(n) => visitor.visit_macro_module(n),
            IRNode::OleObject(n) => visitor.visit_ole_object(n),
            IRNode::ExternalReference(n) => visitor.visit_external_ref(n),
            IRNode::ActiveXControl(n) => visitor.visit_activex_control(n),
            IRNode::Metadata(n) => visitor.visit_metadata(n),
            IRNode::StyleSet(n) => visitor.visit_style_set(n),
            IRNode::NumberingSet(n) => visitor.visit_numbering_set(n),
            IRNode::Comment(n) => visitor.visit_comment(n),
            IRNode::CommentRangeStart(n) => visitor.visit_comment_range_start(n),
            IRNode::CommentRangeEnd(n) => visitor.visit_comment_range_end(n),
            IRNode::CommentReference(n) => visitor.visit_comment_reference(n),
            IRNode::Footnote(n) => visitor.visit_footnote(n),
            IRNode::Endnote(n) => visitor.visit_endnote(n),
            IRNode::Header(n) => visitor.visit_header(n),
            IRNode::Footer(n) => visitor.visit_footer(n),
            IRNode::WordSettings(n) => visitor.visit_word_settings(n),
            IRNode::WebSettings(n) => visitor.visit_web_settings(n),
            IRNode::FontTable(n) => visitor.visit_font_table(n),
            IRNode::ContentControl(n) => visitor.visit_content_control(n),
            IRNode::BookmarkStart(n) => visitor.visit_bookmark_start(n),
            IRNode::BookmarkEnd(n) => visitor.visit_bookmark_end(n),
            IRNode::Field(n) => visitor.visit_field(n),
            IRNode::Revision(n) => visitor.visit_revision(n),
            IRNode::CommentExtensionSet(n) => visitor.visit_comment_extension_set(n),
            IRNode::CommentIdMap(n) => visitor.visit_comment_id_map(n),
            IRNode::SlideMaster(n) => visitor.visit_slide_master(n),
            IRNode::SlideLayout(n) => visitor.visit_slide_layout(n),
            IRNode::NotesMaster(n) => visitor.visit_notes_master(n),
            IRNode::HandoutMaster(n) => visitor.visit_handout_master(n),
            IRNode::NotesSlide(n) => visitor.visit_notes_slide(n),
            IRNode::WorksheetDrawing(n) => visitor.visit_worksheet_drawing(n),
            IRNode::ChartData(n) => visitor.visit_chart_data(n),
            IRNode::PresentationProperties(n) => visitor.visit_presentation_properties(n),
            IRNode::ViewProperties(n) => visitor.visit_view_properties(n),
            IRNode::TableStyleSet(n) => visitor.visit_table_style_set(n),
            IRNode::PptxCommentAuthor(n) => visitor.visit_pptx_comment_author(n),
            IRNode::PptxComment(n) => visitor.visit_pptx_comment(n),
            IRNode::PresentationTag(n) => visitor.visit_presentation_tag(n),
            IRNode::PresentationInfo(n) => visitor.visit_presentation_info(n),
            IRNode::PeoplePart(n) => visitor.visit_people_part(n),
            IRNode::SmartArtPart(n) => visitor.visit_smartart_part(n),
            IRNode::WebExtension(n) => visitor.visit_web_extension(n),
            IRNode::WebExtensionTaskpane(n) => visitor.visit_web_extension_taskpane(n),
            IRNode::GlossaryDocument(n) => visitor.visit_glossary_document(n),
            IRNode::GlossaryEntry(n) => visitor.visit_glossary_entry(n),
            IRNode::VmlDrawing(n) => visitor.visit_vml_drawing(n),
            IRNode::VmlShape(n) => visitor.visit_vml_shape(n),
            IRNode::DrawingPart(n) => visitor.visit_drawing_part(n),
            IRNode::ExternalLinkPart(n) => visitor.visit_external_link_part(n),
            IRNode::ConnectionPart(n) => visitor.visit_connection_part(n),
            IRNode::SlicerPart(n) => visitor.visit_slicer_part(n),
            IRNode::TimelinePart(n) => visitor.visit_timeline_part(n),
            IRNode::QueryTablePart(n) => visitor.visit_query_table_part(n),
            IRNode::Diagnostics(n) => visitor.visit_diagnostics(n),
            IRNode::Theme(n) => visitor.visit_theme(n),
            IRNode::MediaAsset(n) => visitor.visit_media_asset(n),
            IRNode::CustomXmlPart(n) => visitor.visit_custom_xml_part(n),
            IRNode::RelationshipGraph(n) => visitor.visit_relationship_graph(n),
            IRNode::DigitalSignature(n) => visitor.visit_digital_signature(n),
            IRNode::ExtensionPart(n) => visitor.visit_extension_part(n),
        }
    }

    fn get_children(&self, node: &IRNode) -> Vec<NodeId> {
        match node {
            IRNode::Document(n) => n.children(),
            IRNode::Section(n) => n.children(),
            IRNode::Paragraph(n) => n.children(),
            IRNode::Run(_) => vec![],
            IRNode::Hyperlink(n) => n.children(),
            IRNode::Table(n) => n.children(),
            IRNode::TableRow(n) => n.children(),
            IRNode::TableCell(n) => n.children(),
            IRNode::Slide(n) => n.children(),
            IRNode::Shape(n) => n.table.into_iter().collect(),
            IRNode::Worksheet(n) => n.children(),
            IRNode::Cell(_) => vec![],
            IRNode::SharedStringTable(_) => vec![],
            IRNode::SpreadsheetStyles(_) => vec![],
            IRNode::DefinedName(_) => vec![],
            IRNode::ConditionalFormat(_) => vec![],
            IRNode::DataValidation(_) => vec![],
            IRNode::TableDefinition(_) => vec![],
            IRNode::PivotTable(_) => vec![],
            IRNode::PivotCache(_) => vec![],
            IRNode::PivotCacheRecords(_) => vec![],
            IRNode::CalcChain(_) => vec![],
            IRNode::SheetComment(_) => vec![],
            IRNode::SheetMetadata(_) => vec![],
            IRNode::WorkbookProperties(_) => vec![],
            IRNode::MacroProject(n) => n.children(),
            IRNode::MacroModule(_) => vec![],
            IRNode::OleObject(_) => vec![],
            IRNode::ExternalReference(_) => vec![],
            IRNode::ActiveXControl(_) => vec![],
            IRNode::Metadata(_) => vec![],
            IRNode::StyleSet(_) => vec![],
            IRNode::NumberingSet(_) => vec![],
            IRNode::Comment(n) => n.content.clone(),
            IRNode::CommentRangeStart(_) => vec![],
            IRNode::CommentRangeEnd(_) => vec![],
            IRNode::CommentReference(_) => vec![],
            IRNode::Footnote(n) => n.content.clone(),
            IRNode::Endnote(n) => n.content.clone(),
            IRNode::Header(n) => n.content.clone(),
            IRNode::Footer(n) => n.content.clone(),
            IRNode::WordSettings(_) => vec![],
            IRNode::WebSettings(_) => vec![],
            IRNode::FontTable(_) => vec![],
            IRNode::ContentControl(n) => n.content.clone(),
            IRNode::BookmarkStart(_) => vec![],
            IRNode::BookmarkEnd(_) => vec![],
            IRNode::Field(n) => n.runs.clone(),
            IRNode::Revision(n) => n.content.clone(),
            IRNode::CommentExtensionSet(_) => vec![],
            IRNode::CommentIdMap(_) => vec![],
            IRNode::SlideMaster(n) => n.children(),
            IRNode::SlideLayout(n) => n.children(),
            IRNode::NotesMaster(n) => n.children(),
            IRNode::HandoutMaster(n) => n.children(),
            IRNode::NotesSlide(n) => n.shapes.clone(),
            IRNode::WorksheetDrawing(n) => n.children(),
            IRNode::ChartData(_) => vec![],
            IRNode::PresentationProperties(_) => vec![],
            IRNode::ViewProperties(_) => vec![],
            IRNode::TableStyleSet(_) => vec![],
            IRNode::PptxCommentAuthor(_) => vec![],
            IRNode::PptxComment(_) => vec![],
            IRNode::PresentationTag(_) => vec![],
            IRNode::PresentationInfo(_) => vec![],
            IRNode::PeoplePart(_) => vec![],
            IRNode::SmartArtPart(_) => vec![],
            IRNode::WebExtension(_) => vec![],
            IRNode::WebExtensionTaskpane(_) => vec![],
            IRNode::GlossaryDocument(n) => n.entries.clone(),
            IRNode::GlossaryEntry(n) => n.content.clone(),
            IRNode::VmlDrawing(n) => n.shapes.clone(),
            IRNode::VmlShape(_) => vec![],
            IRNode::DrawingPart(n) => n.shapes.clone(),
            IRNode::ExternalLinkPart(_) => vec![],
            IRNode::ConnectionPart(_) => vec![],
            IRNode::SlicerPart(_) => vec![],
            IRNode::TimelinePart(_) => vec![],
            IRNode::QueryTablePart(_) => vec![],
            IRNode::Diagnostics(_) => vec![],
            IRNode::Theme(_) => vec![],
            IRNode::MediaAsset(_) => vec![],
            IRNode::CustomXmlPart(_) => vec![],
            IRNode::RelationshipGraph(_) => vec![],
            IRNode::DigitalSignature(_) => vec![],
            IRNode::ExtensionPart(_) => vec![],
        }
    }
}
