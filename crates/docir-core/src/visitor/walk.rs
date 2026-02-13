use crate::ir::*;
use crate::types::NodeId;

use super::{IrStore, IrVisitor, VisitControl, VisitorResult};

macro_rules! dispatch_visit_node {
    ($node:expr, $visitor:expr, $($variant:ident => $method:ident),+ $(,)?) => {
        match $node {
            $(IRNode::$variant(n) => $visitor.$method(n),)+
        }
    };
}

macro_rules! visit_node_variants {
    ($node:expr, $visitor:expr) => {
        dispatch_visit_node!(
            $node,
            $visitor,
            Document => visit_document,
            Section => visit_section,
            Paragraph => visit_paragraph,
            Run => visit_run,
            Hyperlink => visit_hyperlink,
            Table => visit_table,
            TableRow => visit_table_row,
            TableCell => visit_table_cell,
            Slide => visit_slide,
            Shape => visit_shape,
            Worksheet => visit_worksheet,
            Cell => visit_cell,
            SharedStringTable => visit_shared_string_table,
            SpreadsheetStyles => visit_spreadsheet_styles,
            DefinedName => visit_defined_name,
            ConditionalFormat => visit_conditional_format,
            DataValidation => visit_data_validation,
            TableDefinition => visit_table_definition,
            PivotTable => visit_pivot_table,
            PivotCache => visit_pivot_cache,
            PivotCacheRecords => visit_pivot_cache_records,
            CalcChain => visit_calc_chain,
            SheetComment => visit_sheet_comment,
            SheetMetadata => visit_sheet_metadata,
            WorkbookProperties => visit_workbook_properties,
            MacroProject => visit_macro_project,
            MacroModule => visit_macro_module,
            OleObject => visit_ole_object,
            ExternalReference => visit_external_ref,
            ActiveXControl => visit_activex_control,
            Metadata => visit_metadata,
            StyleSet => visit_style_set,
            NumberingSet => visit_numbering_set,
            Comment => visit_comment,
            CommentRangeStart => visit_comment_range_start,
            CommentRangeEnd => visit_comment_range_end,
            CommentReference => visit_comment_reference,
            Footnote => visit_footnote,
            Endnote => visit_endnote,
            Header => visit_header,
            Footer => visit_footer,
            WordSettings => visit_word_settings,
            WebSettings => visit_web_settings,
            FontTable => visit_font_table,
            ContentControl => visit_content_control,
            BookmarkStart => visit_bookmark_start,
            BookmarkEnd => visit_bookmark_end,
            Field => visit_field,
            Revision => visit_revision,
            CommentExtensionSet => visit_comment_extension_set,
            CommentIdMap => visit_comment_id_map,
            SlideMaster => visit_slide_master,
            SlideLayout => visit_slide_layout,
            NotesMaster => visit_notes_master,
            HandoutMaster => visit_handout_master,
            NotesSlide => visit_notes_slide,
            WorksheetDrawing => visit_worksheet_drawing,
            ChartData => visit_chart_data,
            PresentationProperties => visit_presentation_properties,
            ViewProperties => visit_view_properties,
            TableStyleSet => visit_table_style_set,
            PptxCommentAuthor => visit_pptx_comment_author,
            PptxComment => visit_pptx_comment,
            PresentationTag => visit_presentation_tag,
            PresentationInfo => visit_presentation_info,
            PeoplePart => visit_people_part,
            SmartArtPart => visit_smartart_part,
            WebExtension => visit_web_extension,
            WebExtensionTaskpane => visit_web_extension_taskpane,
            GlossaryDocument => visit_glossary_document,
            GlossaryEntry => visit_glossary_entry,
            VmlDrawing => visit_vml_drawing,
            VmlShape => visit_vml_shape,
            DrawingPart => visit_drawing_part,
            ExternalLinkPart => visit_external_link_part,
            ConnectionPart => visit_connection_part,
            SlicerPart => visit_slicer_part,
            TimelinePart => visit_timeline_part,
            QueryTablePart => visit_query_table_part,
            Diagnostics => visit_diagnostics,
            Theme => visit_theme,
            MediaAsset => visit_media_asset,
            CustomXmlPart => visit_custom_xml_part,
            RelationshipGraph => visit_relationship_graph,
            DigitalSignature => visit_digital_signature,
            ExtensionPart => visit_extension_part
        )
    };
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
        visit_node_variants!(node, visitor)
    }

    fn get_children(&self, node: &IRNode) -> Vec<NodeId> {
        node.children()
    }
}
