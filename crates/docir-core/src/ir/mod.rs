//! Intermediate Representation (IR) node definitions.
//!
//! This module contains the core IR nodes that represent the semantic
//! structure of Office documents.

mod addins;
mod builder;
mod charts;
mod connections;
mod custom_xml;
mod diagnostics;
mod document;
mod drawingml;
mod external_links;
mod glossary;
mod media;
mod metadata;
pub(crate) mod node_list;
mod notes;
mod numbering;
mod package;
mod paragraph;
mod presentation;
mod query_tables;
mod signature;
mod slicers;
mod spreadsheet;
mod style;
mod table;
mod theme;
mod vml;
mod word_controls;
mod word_revisions;
mod word_settings;

pub use addins::*;
pub use builder::*;
pub use charts::*;
pub use connections::*;
pub use custom_xml::*;
pub use diagnostics::*;
pub use document::*;
pub use drawingml::*;
pub use external_links::*;
pub use glossary::*;
pub use media::*;
pub use metadata::*;
pub use notes::*;
pub use numbering::*;
pub use package::*;
pub use paragraph::*;
pub use presentation::*;
pub use query_tables::*;
pub use signature::*;
pub use slicers::*;
pub use spreadsheet::*;
pub use style::*;
pub use table::*;
pub use theme::*;
pub use vml::*;
pub use word_controls::*;
pub use word_revisions::*;
pub use word_settings::*;

use crate::security::{ActiveXControl, ExternalReference, MacroModule, MacroProject, OleObject};
use crate::types::{NodeId, NodeType, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

pub(crate) fn new_node_id() -> NodeId {
    NodeId::new()
}

/// Trait implemented by all IR nodes.
pub trait IrNode {
    /// Returns the unique ID of this node.
    fn node_id(&self) -> NodeId;

    /// Returns the type of this node.
    fn node_type(&self) -> NodeType;

    /// Returns the IDs of child nodes.
    fn children(&self) -> Vec<NodeId>;

    /// Returns the source span, if available.
    fn source_span(&self) -> Option<&SourceSpan>;
}

/// Enumeration of all possible IR nodes.
///
/// This is the main type used for working with the IR tree.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
#[serde(tag = "type")]
pub enum IRNode {
    Document(Document),
    Section(Section),
    Paragraph(Paragraph),
    Run(Run),
    Hyperlink(Hyperlink),
    Table(Table),
    TableRow(TableRow),
    TableCell(TableCell),
    Slide(Slide),
    Shape(Shape),
    Worksheet(Worksheet),
    Cell(Cell),
    SharedStringTable(SharedStringTable),
    SpreadsheetStyles(SpreadsheetStyles),
    DefinedName(DefinedName),
    ConditionalFormat(ConditionalFormat),
    DataValidation(DataValidation),
    TableDefinition(TableDefinition),
    PivotTable(PivotTable),
    PivotCache(PivotCache),
    PivotCacheRecords(PivotCacheRecords),
    CalcChain(CalcChain),
    SheetComment(SheetComment),
    SheetMetadata(SheetMetadata),
    WorkbookProperties(WorkbookProperties),
    MacroProject(MacroProject),
    MacroModule(MacroModule),
    OleObject(OleObject),
    ExternalReference(ExternalReference),
    ActiveXControl(ActiveXControl),
    Metadata(DocumentMetadata),
    StyleSet(StyleSet),
    NumberingSet(NumberingSet),
    Comment(Comment),
    CommentRangeStart(CommentRangeStart),
    CommentRangeEnd(CommentRangeEnd),
    CommentReference(CommentReference),
    Footnote(Footnote),
    Endnote(Endnote),
    Header(Header),
    Footer(Footer),
    WordSettings(WordSettings),
    WebSettings(WebSettings),
    FontTable(FontTable),
    ContentControl(ContentControl),
    BookmarkStart(BookmarkStart),
    BookmarkEnd(BookmarkEnd),
    Field(Field),
    Revision(Revision),
    CommentExtensionSet(CommentExtensionSet),
    CommentIdMap(CommentIdMap),
    SlideMaster(SlideMaster),
    SlideLayout(SlideLayout),
    NotesMaster(NotesMaster),
    HandoutMaster(HandoutMaster),
    NotesSlide(NotesSlide),
    WorksheetDrawing(WorksheetDrawing),
    ChartData(ChartData),
    PresentationProperties(PresentationProperties),
    ViewProperties(ViewProperties),
    TableStyleSet(TableStyleSet),
    PptxCommentAuthor(PptxCommentAuthor),
    PptxComment(PptxComment),
    PresentationTag(PresentationTag),
    PresentationInfo(PresentationInfo),
    PeoplePart(PeoplePart),
    SmartArtPart(SmartArtPart),
    WebExtension(WebExtension),
    WebExtensionTaskpane(WebExtensionTaskpane),
    GlossaryDocument(GlossaryDocument),
    GlossaryEntry(GlossaryEntry),
    VmlDrawing(VmlDrawing),
    VmlShape(VmlShape),
    DrawingPart(DrawingPart),
    ExternalLinkPart(ExternalLinkPart),
    ConnectionPart(ConnectionPart),
    SlicerPart(SlicerPart),
    TimelinePart(TimelinePart),
    QueryTablePart(QueryTablePart),
    Diagnostics(Diagnostics),
    Theme(Theme),
    MediaAsset(MediaAsset),
    CustomXmlPart(CustomXmlPart),
    RelationshipGraph(RelationshipGraph),
    DigitalSignature(DigitalSignature),
    ExtensionPart(ExtensionPart),
}

macro_rules! for_each_ir_node_variant {
    ($macro:ident) => {
        $macro!(
            Document,
            Section,
            Paragraph,
            Run,
            Hyperlink,
            Table,
            TableRow,
            TableCell,
            Slide,
            Shape,
            Worksheet,
            Cell,
            SharedStringTable,
            SpreadsheetStyles,
            DefinedName,
            ConditionalFormat,
            DataValidation,
            TableDefinition,
            PivotTable,
            PivotCache,
            PivotCacheRecords,
            CalcChain,
            SheetComment,
            SheetMetadata,
            WorkbookProperties,
            MacroProject,
            MacroModule,
            OleObject,
            ExternalReference,
            ActiveXControl,
            Metadata,
            StyleSet,
            NumberingSet,
            Comment,
            CommentRangeStart,
            CommentRangeEnd,
            CommentReference,
            Footnote,
            Endnote,
            Header,
            Footer,
            WordSettings,
            WebSettings,
            FontTable,
            ContentControl,
            BookmarkStart,
            BookmarkEnd,
            Field,
            Revision,
            CommentExtensionSet,
            CommentIdMap,
            SlideMaster,
            SlideLayout,
            NotesMaster,
            HandoutMaster,
            NotesSlide,
            WorksheetDrawing,
            ChartData,
            PresentationProperties,
            ViewProperties,
            TableStyleSet,
            PptxCommentAuthor,
            PptxComment,
            PresentationTag,
            PresentationInfo,
            PeoplePart,
            SmartArtPart,
            WebExtension,
            WebExtensionTaskpane,
            GlossaryDocument,
            GlossaryEntry,
            VmlDrawing,
            VmlShape,
            DrawingPart,
            ExternalLinkPart,
            ConnectionPart,
            SlicerPart,
            TimelinePart,
            QueryTablePart,
            Diagnostics,
            Theme,
            MediaAsset,
            CustomXmlPart,
            RelationshipGraph,
            DigitalSignature,
            ExtensionPart
        )
    };
}

macro_rules! for_each_ir_span_variant {
    ($macro:ident) => {
        $macro!(
            Document,
            Section,
            Paragraph,
            Run,
            Hyperlink,
            Table,
            TableRow,
            TableCell,
            Slide,
            Shape,
            Worksheet,
            Cell,
            SharedStringTable,
            SpreadsheetStyles,
            DefinedName,
            ConditionalFormat,
            DataValidation,
            TableDefinition,
            PivotTable,
            PivotCache,
            PivotCacheRecords,
            CalcChain,
            SheetComment,
            SheetMetadata,
            WorkbookProperties,
            MacroProject,
            MacroModule,
            OleObject,
            ExternalReference,
            ActiveXControl,
            StyleSet,
            NumberingSet,
            Comment,
            CommentRangeStart,
            CommentRangeEnd,
            CommentReference,
            Footnote,
            Endnote,
            Header,
            Footer,
            WordSettings,
            WebSettings,
            FontTable,
            ContentControl,
            BookmarkStart,
            BookmarkEnd,
            Field,
            Revision,
            CommentExtensionSet,
            CommentIdMap,
            SlideMaster,
            SlideLayout,
            NotesMaster,
            HandoutMaster,
            NotesSlide,
            WorksheetDrawing,
            ChartData,
            PresentationProperties,
            ViewProperties,
            TableStyleSet,
            PptxCommentAuthor,
            PptxComment,
            PresentationTag,
            PresentationInfo,
            PeoplePart,
            SmartArtPart,
            WebExtension,
            WebExtensionTaskpane,
            GlossaryDocument,
            GlossaryEntry,
            VmlDrawing,
            VmlShape,
            DrawingPart,
            ExternalLinkPart,
            ConnectionPart,
            SlicerPart,
            TimelinePart,
            QueryTablePart,
            Diagnostics,
            Theme,
            MediaAsset,
            CustomXmlPart,
            RelationshipGraph,
            DigitalSignature,
            ExtensionPart
        )
    };
}

macro_rules! node_id_match {
    ($value:expr; $($variant:ident),+ $(,)?) => {
        match $value {
            $(IRNode::$variant(n) => n.id,)+
        }
    };
}

macro_rules! node_type_match {
    ($value:expr; $($variant:ident),+ $(,)?) => {
        match $value {
            $(IRNode::$variant(_) => NodeType::$variant,)+
        }
    };
}

impl IrNode for IRNode {
    fn node_id(&self) -> NodeId {
        macro_rules! node_id_arms {
            ($($variant:ident),+ $(,)?) => {
                node_id_match!(self; $($variant),+)
            };
        }
        for_each_ir_node_variant!(node_id_arms)
    }

    fn node_type(&self) -> NodeType {
        macro_rules! node_type_arms {
            ($($variant:ident),+ $(,)?) => {
                node_type_match!(self; $($variant),+)
            };
        }
        for_each_ir_node_variant!(node_type_arms)
    }

    fn children(&self) -> Vec<NodeId> {
        match self {
            IRNode::Document(n) => n.children(),
            IRNode::Section(n) => n.children(),
            IRNode::Paragraph(n) => n.children(),
            IRNode::Hyperlink(n) => n.children(),
            IRNode::Table(n) => n.children(),
            IRNode::TableRow(n) => n.children(),
            IRNode::TableCell(n) => n.children(),
            IRNode::Slide(n) => n.children(),
            IRNode::Shape(n) => n.table.into_iter().collect(),
            IRNode::Worksheet(n) => n.children(),
            IRNode::MacroProject(n) => n.children(),
            IRNode::Comment(n) => n.content.clone(),
            IRNode::Footnote(n) => n.content.clone(),
            IRNode::Endnote(n) => n.content.clone(),
            IRNode::Header(n) => n.content.clone(),
            IRNode::Footer(n) => n.content.clone(),
            IRNode::ContentControl(n) => n.content.clone(),
            IRNode::Field(n) => n.runs.clone(),
            IRNode::Revision(n) => n.content.clone(),
            IRNode::SlideMaster(n) => n.children(),
            IRNode::SlideLayout(n) => n.children(),
            IRNode::NotesMaster(n) => n.children(),
            IRNode::HandoutMaster(n) => n.children(),
            IRNode::NotesSlide(n) => n.shapes.clone(),
            IRNode::WorksheetDrawing(n) => n.children(),
            IRNode::GlossaryDocument(n) => n.entries.clone(),
            IRNode::GlossaryEntry(n) => n.content.clone(),
            IRNode::VmlDrawing(n) => n.shapes.clone(),
            IRNode::DrawingPart(n) => n.shapes.clone(),
            _ => Vec::new(),
        }
    }

    fn source_span(&self) -> Option<&SourceSpan> {
        macro_rules! source_span_arms {
            ($($variant:ident),+ $(,)?) => {
                match self {
                    $(IRNode::$variant(n) => n.span.as_ref(),)+
                    IRNode::Metadata(_) => None,
                }
            };
        }
        for_each_ir_span_variant!(source_span_arms)
    }
}
