//! Visitor pattern for IR traversal.
//!
//! This module provides traits and utilities for traversing the IR tree
//! in various orders (pre-order, post-order) and performing operations
//! on nodes.

use crate::error::CoreError;
use crate::ir::*;
use crate::security::*;
use crate::types::NodeId;
use std::collections::HashMap;

/// Result type for visitor operations.
pub type VisitorResult<T> = Result<T, CoreError>;

/// Control flow for visitor traversal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VisitControl {
    /// Continue traversal normally.
    Continue,
    /// Skip children of current node.
    SkipChildren,
    /// Stop traversal entirely.
    Stop,
}

/// Trait for immutable IR traversal.
///
/// Implement this trait to perform read-only operations on the IR tree.
/// Default implementations return `Continue` for all node types.
pub trait IrVisitor {
    /// Called when entering a Document node.
    fn visit_document(&mut self, _doc: &Document) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Section node.
    fn visit_section(&mut self, _section: &Section) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Paragraph node.
    fn visit_paragraph(&mut self, _para: &Paragraph) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Run node.
    fn visit_run(&mut self, _run: &Run) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Hyperlink node.
    fn visit_hyperlink(&mut self, _link: &Hyperlink) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Table node.
    fn visit_table(&mut self, _table: &Table) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a TableRow node.
    fn visit_table_row(&mut self, _row: &TableRow) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a TableCell node.
    fn visit_table_cell(&mut self, _cell: &TableCell) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Slide node.
    fn visit_slide(&mut self, _slide: &Slide) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Shape node.
    fn visit_shape(&mut self, _shape: &Shape) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Worksheet node.
    fn visit_worksheet(&mut self, _sheet: &Worksheet) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Cell node.
    fn visit_cell(&mut self, _cell: &Cell) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a SharedStringTable node.
    fn visit_shared_string_table(
        &mut self,
        _table: &SharedStringTable,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a SpreadsheetStyles node.
    fn visit_spreadsheet_styles(
        &mut self,
        _styles: &SpreadsheetStyles,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a DefinedName node.
    fn visit_defined_name(&mut self, _name: &DefinedName) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a ConditionalFormat node.
    fn visit_conditional_format(
        &mut self,
        _format: &ConditionalFormat,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a DataValidation node.
    fn visit_data_validation(
        &mut self,
        _validation: &DataValidation,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a TableDefinition node.
    fn visit_table_definition(&mut self, _table: &TableDefinition) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a PivotTable node.
    fn visit_pivot_table(&mut self, _pivot: &PivotTable) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a PivotCache node.
    fn visit_pivot_cache(&mut self, _cache: &PivotCache) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a PivotCacheRecords node.
    fn visit_pivot_cache_records(
        &mut self,
        _records: &PivotCacheRecords,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a CalcChain node.
    fn visit_calc_chain(&mut self, _chain: &CalcChain) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a SheetComment node.
    fn visit_sheet_comment(&mut self, _comment: &SheetComment) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a SheetMetadata node.
    fn visit_sheet_metadata(&mut self, _meta: &SheetMetadata) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a WorkbookProperties node.
    fn visit_workbook_properties(
        &mut self,
        _props: &WorkbookProperties,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a MacroProject node.
    fn visit_macro_project(&mut self, _project: &MacroProject) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a MacroModule node.
    fn visit_macro_module(&mut self, _module: &MacroModule) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering an OleObject node.
    fn visit_ole_object(&mut self, _ole: &OleObject) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering an ExternalReference node.
    fn visit_external_ref(&mut self, _ext_ref: &ExternalReference) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering an ActiveXControl node.
    fn visit_activex_control(&mut self, _control: &ActiveXControl) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Metadata node.
    fn visit_metadata(&mut self, _meta: &DocumentMetadata) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a StyleSet node.
    fn visit_style_set(&mut self, _styles: &StyleSet) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a NumberingSet node.
    fn visit_numbering_set(&mut self, _num: &NumberingSet) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Comment node.
    fn visit_comment(&mut self, _comment: &Comment) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a CommentRangeStart node.
    fn visit_comment_range_start(
        &mut self,
        _range: &CommentRangeStart,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a CommentRangeEnd node.
    fn visit_comment_range_end(&mut self, _range: &CommentRangeEnd) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a CommentReference node.
    fn visit_comment_reference(
        &mut self,
        _reference: &CommentReference,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Footnote node.
    fn visit_footnote(&mut self, _note: &Footnote) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering an Endnote node.
    fn visit_endnote(&mut self, _note: &Endnote) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Header node.
    fn visit_header(&mut self, _header: &Header) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Footer node.
    fn visit_footer(&mut self, _footer: &Footer) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a WordSettings node.
    fn visit_word_settings(&mut self, _settings: &WordSettings) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a WebSettings node.
    fn visit_web_settings(&mut self, _settings: &WebSettings) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a FontTable node.
    fn visit_font_table(&mut self, _table: &FontTable) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a ContentControl node.
    fn visit_content_control(&mut self, _control: &ContentControl) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a BookmarkStart node.
    fn visit_bookmark_start(&mut self, _bookmark: &BookmarkStart) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a BookmarkEnd node.
    fn visit_bookmark_end(&mut self, _bookmark: &BookmarkEnd) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Field node.
    fn visit_field(&mut self, _field: &Field) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Revision node.
    fn visit_revision(&mut self, _revision: &Revision) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a CommentExtensionSet node.
    fn visit_comment_extension_set(
        &mut self,
        _set: &CommentExtensionSet,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a CommentIdMap node.
    fn visit_comment_id_map(&mut self, _map: &CommentIdMap) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a SlideMaster node.
    fn visit_slide_master(&mut self, _master: &SlideMaster) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a SlideLayout node.
    fn visit_slide_layout(&mut self, _layout: &SlideLayout) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a NotesMaster node.
    fn visit_notes_master(&mut self, _master: &NotesMaster) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a HandoutMaster node.
    fn visit_handout_master(&mut self, _master: &HandoutMaster) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a NotesSlide node.
    fn visit_notes_slide(&mut self, _slide: &NotesSlide) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a WorksheetDrawing node.
    fn visit_worksheet_drawing(
        &mut self,
        _drawing: &WorksheetDrawing,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a ChartData node.
    fn visit_chart_data(&mut self, _chart: &ChartData) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a PresentationProperties node.
    fn visit_presentation_properties(
        &mut self,
        _props: &PresentationProperties,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a ViewProperties node.
    fn visit_view_properties(&mut self, _props: &ViewProperties) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a TableStyleSet node.
    fn visit_table_style_set(&mut self, _styles: &TableStyleSet) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a PptxCommentAuthor node.
    fn visit_pptx_comment_author(
        &mut self,
        _author: &PptxCommentAuthor,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a PptxComment node.
    fn visit_pptx_comment(&mut self, _comment: &PptxComment) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a PresentationTag node.
    fn visit_presentation_tag(&mut self, _tag: &PresentationTag) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a PresentationInfo node.
    fn visit_presentation_info(&mut self, _info: &PresentationInfo) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a PeoplePart node.
    fn visit_people_part(&mut self, _part: &PeoplePart) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a SmartArtPart node.
    fn visit_smartart_part(&mut self, _part: &SmartArtPart) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a WebExtension node.
    fn visit_web_extension(&mut self, _ext: &WebExtension) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a WebExtensionTaskpane node.
    fn visit_web_extension_taskpane(
        &mut self,
        _pane: &WebExtensionTaskpane,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a GlossaryDocument node.
    fn visit_glossary_document(&mut self, _doc: &GlossaryDocument) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a GlossaryEntry node.
    fn visit_glossary_entry(&mut self, _entry: &GlossaryEntry) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a VmlDrawing node.
    fn visit_vml_drawing(&mut self, _drawing: &VmlDrawing) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a VmlShape node.
    fn visit_vml_shape(&mut self, _shape: &VmlShape) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a DrawingPart node.
    fn visit_drawing_part(&mut self, _part: &DrawingPart) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering an ExternalLinkPart node.
    fn visit_external_link_part(
        &mut self,
        _part: &ExternalLinkPart,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a ConnectionPart node.
    fn visit_connection_part(&mut self, _part: &ConnectionPart) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a SlicerPart node.
    fn visit_slicer_part(&mut self, _part: &SlicerPart) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a TimelinePart node.
    fn visit_timeline_part(&mut self, _part: &TimelinePart) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a QueryTablePart node.
    fn visit_query_table_part(&mut self, _part: &QueryTablePart) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Diagnostics node.
    fn visit_diagnostics(&mut self, _diag: &Diagnostics) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a Theme node.
    fn visit_theme(&mut self, _theme: &Theme) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a MediaAsset node.
    fn visit_media_asset(&mut self, _media: &MediaAsset) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a CustomXmlPart node.
    fn visit_custom_xml_part(&mut self, _part: &CustomXmlPart) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a RelationshipGraph node.
    fn visit_relationship_graph(
        &mut self,
        _graph: &RelationshipGraph,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering a DigitalSignature node.
    fn visit_digital_signature(
        &mut self,
        _sig: &crate::ir::DigitalSignature,
    ) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }

    /// Called when entering an ExtensionPart node.
    fn visit_extension_part(&mut self, _part: &ExtensionPart) -> VisitorResult<VisitControl> {
        Ok(VisitControl::Continue)
    }
}

/// Storage for IR nodes indexed by NodeId.
#[derive(Debug, Clone, Default)]
pub struct IrStore {
    nodes: HashMap<NodeId, IRNode>,
}

impl IrStore {
    /// Creates a new empty IrStore.
    pub fn new() -> Self {
        Self {
            nodes: HashMap::new(),
        }
    }

    /// Consumes the store and returns all nodes.
    pub fn into_nodes(self) -> Vec<IRNode> {
        self.nodes.into_values().collect()
    }

    /// Inserts a node into the store.
    pub fn insert(&mut self, node: IRNode) {
        let id = node.node_id();
        self.nodes.insert(id, node);
    }

    /// Gets a node by ID.
    pub fn get(&self, id: NodeId) -> Option<&IRNode> {
        self.nodes.get(&id)
    }

    /// Gets a mutable reference to a node by ID.
    pub fn get_mut(&mut self, id: NodeId) -> Option<&mut IRNode> {
        self.nodes.get_mut(&id)
    }

    /// Returns the number of nodes.
    pub fn len(&self) -> usize {
        self.nodes.len()
    }

    /// Returns node IDs by type.
    pub fn iter_ids_by_type(
        &self,
        node_type: crate::types::NodeType,
    ) -> impl Iterator<Item = NodeId> + '_ {
        self.nodes.iter().filter_map(move |(id, node)| {
            if node.node_type() == node_type {
                Some(*id)
            } else {
                None
            }
        })
    }

    /// Returns an iterator over all nodes.
    pub fn values(&self) -> impl Iterator<Item = &IRNode> + '_ {
        self.nodes.values()
    }

    /// Returns true if the store is empty.
    pub fn is_empty(&self) -> bool {
        self.nodes.is_empty()
    }

    /// Extends the store with nodes from another store.
    pub fn extend(&mut self, other: IrStore) {
        for (id, node) in other.nodes {
            self.nodes.insert(id, node);
        }
    }

    /// Iterates over all nodes.
    pub fn iter(&self) -> impl Iterator<Item = (&NodeId, &IRNode)> {
        self.nodes.iter()
    }
}

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
