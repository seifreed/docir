//! Intermediate Representation (IR) node definitions.
//!
//! This module contains the core IR nodes that represent the semantic
//! structure of Office documents.

mod addins;
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

use crate::types::{NodeId, NodeType, SourceSpan};
use serde::{Deserialize, Serialize};

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
#[derive(Debug, Clone, Serialize, Deserialize)]
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
    MacroProject(crate::security::MacroProject),
    MacroModule(crate::security::MacroModule),
    OleObject(crate::security::OleObject),
    ExternalReference(crate::security::ExternalReference),
    ActiveXControl(crate::security::ActiveXControl),
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

impl IrNode for IRNode {
    fn node_id(&self) -> NodeId {
        match self {
            IRNode::Document(n) => n.id,
            IRNode::Section(n) => n.id,
            IRNode::Paragraph(n) => n.id,
            IRNode::Run(n) => n.id,
            IRNode::Hyperlink(n) => n.id,
            IRNode::Table(n) => n.id,
            IRNode::TableRow(n) => n.id,
            IRNode::TableCell(n) => n.id,
            IRNode::Slide(n) => n.id,
            IRNode::Shape(n) => n.id,
            IRNode::Worksheet(n) => n.id,
            IRNode::Cell(n) => n.id,
            IRNode::SharedStringTable(n) => n.id,
            IRNode::SpreadsheetStyles(n) => n.id,
            IRNode::DefinedName(n) => n.id,
            IRNode::ConditionalFormat(n) => n.id,
            IRNode::DataValidation(n) => n.id,
            IRNode::TableDefinition(n) => n.id,
            IRNode::PivotTable(n) => n.id,
            IRNode::PivotCache(n) => n.id,
            IRNode::PivotCacheRecords(n) => n.id,
            IRNode::CalcChain(n) => n.id,
            IRNode::SheetComment(n) => n.id,
            IRNode::SheetMetadata(n) => n.id,
            IRNode::WorkbookProperties(n) => n.id,
            IRNode::MacroProject(n) => n.id,
            IRNode::MacroModule(n) => n.id,
            IRNode::OleObject(n) => n.id,
            IRNode::ExternalReference(n) => n.id,
            IRNode::ActiveXControl(n) => n.id,
            IRNode::Metadata(n) => n.id,
            IRNode::StyleSet(n) => n.id,
            IRNode::NumberingSet(n) => n.id,
            IRNode::Comment(n) => n.id,
            IRNode::CommentRangeStart(n) => n.id,
            IRNode::CommentRangeEnd(n) => n.id,
            IRNode::CommentReference(n) => n.id,
            IRNode::Footnote(n) => n.id,
            IRNode::Endnote(n) => n.id,
            IRNode::Header(n) => n.id,
            IRNode::Footer(n) => n.id,
            IRNode::WordSettings(n) => n.id,
            IRNode::WebSettings(n) => n.id,
            IRNode::FontTable(n) => n.id,
            IRNode::ContentControl(n) => n.id,
            IRNode::BookmarkStart(n) => n.id,
            IRNode::BookmarkEnd(n) => n.id,
            IRNode::Field(n) => n.id,
            IRNode::Revision(n) => n.id,
            IRNode::CommentExtensionSet(n) => n.id,
            IRNode::CommentIdMap(n) => n.id,
            IRNode::SlideMaster(n) => n.id,
            IRNode::SlideLayout(n) => n.id,
            IRNode::NotesMaster(n) => n.id,
            IRNode::HandoutMaster(n) => n.id,
            IRNode::NotesSlide(n) => n.id,
            IRNode::WorksheetDrawing(n) => n.id,
            IRNode::ChartData(n) => n.id,
            IRNode::PresentationProperties(n) => n.id,
            IRNode::ViewProperties(n) => n.id,
            IRNode::TableStyleSet(n) => n.id,
            IRNode::PptxCommentAuthor(n) => n.id,
            IRNode::PptxComment(n) => n.id,
            IRNode::PresentationTag(n) => n.id,
            IRNode::PresentationInfo(n) => n.id,
            IRNode::PeoplePart(n) => n.id,
            IRNode::SmartArtPart(n) => n.id,
            IRNode::WebExtension(n) => n.id,
            IRNode::WebExtensionTaskpane(n) => n.id,
            IRNode::GlossaryDocument(n) => n.id,
            IRNode::GlossaryEntry(n) => n.id,
            IRNode::VmlDrawing(n) => n.id,
            IRNode::VmlShape(n) => n.id,
            IRNode::DrawingPart(n) => n.id,
            IRNode::ExternalLinkPart(n) => n.id,
            IRNode::ConnectionPart(n) => n.id,
            IRNode::SlicerPart(n) => n.id,
            IRNode::TimelinePart(n) => n.id,
            IRNode::QueryTablePart(n) => n.id,
            IRNode::Diagnostics(n) => n.id,
            IRNode::Theme(n) => n.id,
            IRNode::MediaAsset(n) => n.id,
            IRNode::CustomXmlPart(n) => n.id,
            IRNode::RelationshipGraph(n) => n.id,
            IRNode::DigitalSignature(n) => n.id,
            IRNode::ExtensionPart(n) => n.id,
        }
    }

    fn node_type(&self) -> NodeType {
        match self {
            IRNode::Document(_) => NodeType::Document,
            IRNode::Section(_) => NodeType::Section,
            IRNode::Paragraph(_) => NodeType::Paragraph,
            IRNode::Run(_) => NodeType::Run,
            IRNode::Hyperlink(_) => NodeType::Hyperlink,
            IRNode::Table(_) => NodeType::Table,
            IRNode::TableRow(_) => NodeType::TableRow,
            IRNode::TableCell(_) => NodeType::TableCell,
            IRNode::Slide(_) => NodeType::Slide,
            IRNode::Shape(_) => NodeType::Shape,
            IRNode::Worksheet(_) => NodeType::Worksheet,
            IRNode::Cell(_) => NodeType::Cell,
            IRNode::SharedStringTable(_) => NodeType::SharedStringTable,
            IRNode::SpreadsheetStyles(_) => NodeType::SpreadsheetStyles,
            IRNode::DefinedName(_) => NodeType::DefinedName,
            IRNode::ConditionalFormat(_) => NodeType::ConditionalFormat,
            IRNode::DataValidation(_) => NodeType::DataValidation,
            IRNode::TableDefinition(_) => NodeType::TableDefinition,
            IRNode::PivotTable(_) => NodeType::PivotTable,
            IRNode::PivotCache(_) => NodeType::PivotCache,
            IRNode::PivotCacheRecords(_) => NodeType::PivotCacheRecords,
            IRNode::CalcChain(_) => NodeType::CalcChain,
            IRNode::SheetComment(_) => NodeType::SheetComment,
            IRNode::SheetMetadata(_) => NodeType::SheetMetadata,
            IRNode::WorkbookProperties(_) => NodeType::WorkbookProperties,
            IRNode::MacroProject(_) => NodeType::MacroProject,
            IRNode::MacroModule(_) => NodeType::MacroModule,
            IRNode::OleObject(_) => NodeType::OleObject,
            IRNode::ExternalReference(_) => NodeType::ExternalReference,
            IRNode::ActiveXControl(_) => NodeType::ActiveXControl,
            IRNode::Metadata(_) => NodeType::Metadata,
            IRNode::StyleSet(_) => NodeType::StyleSet,
            IRNode::NumberingSet(_) => NodeType::NumberingSet,
            IRNode::Comment(_) => NodeType::Comment,
            IRNode::CommentRangeStart(_) => NodeType::CommentRangeStart,
            IRNode::CommentRangeEnd(_) => NodeType::CommentRangeEnd,
            IRNode::CommentReference(_) => NodeType::CommentReference,
            IRNode::Footnote(_) => NodeType::Footnote,
            IRNode::Endnote(_) => NodeType::Endnote,
            IRNode::Header(_) => NodeType::Header,
            IRNode::Footer(_) => NodeType::Footer,
            IRNode::WordSettings(_) => NodeType::WordSettings,
            IRNode::WebSettings(_) => NodeType::WebSettings,
            IRNode::FontTable(_) => NodeType::FontTable,
            IRNode::ContentControl(_) => NodeType::ContentControl,
            IRNode::BookmarkStart(_) => NodeType::BookmarkStart,
            IRNode::BookmarkEnd(_) => NodeType::BookmarkEnd,
            IRNode::Field(_) => NodeType::Field,
            IRNode::Revision(_) => NodeType::Revision,
            IRNode::CommentExtensionSet(_) => NodeType::CommentExtensionSet,
            IRNode::CommentIdMap(_) => NodeType::CommentIdMap,
            IRNode::SlideMaster(_) => NodeType::SlideMaster,
            IRNode::SlideLayout(_) => NodeType::SlideLayout,
            IRNode::NotesMaster(_) => NodeType::NotesMaster,
            IRNode::HandoutMaster(_) => NodeType::HandoutMaster,
            IRNode::NotesSlide(_) => NodeType::NotesSlide,
            IRNode::WorksheetDrawing(_) => NodeType::WorksheetDrawing,
            IRNode::ChartData(_) => NodeType::ChartData,
            IRNode::PresentationProperties(_) => NodeType::PresentationProperties,
            IRNode::ViewProperties(_) => NodeType::ViewProperties,
            IRNode::TableStyleSet(_) => NodeType::TableStyleSet,
            IRNode::PptxCommentAuthor(_) => NodeType::PptxCommentAuthor,
            IRNode::PptxComment(_) => NodeType::PptxComment,
            IRNode::PresentationTag(_) => NodeType::PresentationTag,
            IRNode::PresentationInfo(_) => NodeType::PresentationInfo,
            IRNode::PeoplePart(_) => NodeType::PeoplePart,
            IRNode::SmartArtPart(_) => NodeType::SmartArtPart,
            IRNode::WebExtension(_) => NodeType::WebExtension,
            IRNode::WebExtensionTaskpane(_) => NodeType::WebExtensionTaskpane,
            IRNode::GlossaryDocument(_) => NodeType::GlossaryDocument,
            IRNode::GlossaryEntry(_) => NodeType::GlossaryEntry,
            IRNode::VmlDrawing(_) => NodeType::VmlDrawing,
            IRNode::VmlShape(_) => NodeType::VmlShape,
            IRNode::DrawingPart(_) => NodeType::DrawingPart,
            IRNode::ExternalLinkPart(_) => NodeType::ExternalLinkPart,
            IRNode::ConnectionPart(_) => NodeType::ConnectionPart,
            IRNode::SlicerPart(_) => NodeType::SlicerPart,
            IRNode::TimelinePart(_) => NodeType::TimelinePart,
            IRNode::QueryTablePart(_) => NodeType::QueryTablePart,
            IRNode::Diagnostics(_) => NodeType::Diagnostics,
            IRNode::Theme(_) => NodeType::Theme,
            IRNode::MediaAsset(_) => NodeType::MediaAsset,
            IRNode::CustomXmlPart(_) => NodeType::CustomXmlPart,
            IRNode::RelationshipGraph(_) => NodeType::RelationshipGraph,
            IRNode::DigitalSignature(_) => NodeType::DigitalSignature,
            IRNode::ExtensionPart(_) => NodeType::ExtensionPart,
        }
    }

    fn children(&self) -> Vec<NodeId> {
        match self {
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

    fn source_span(&self) -> Option<&SourceSpan> {
        match self {
            IRNode::Document(n) => n.span.as_ref(),
            IRNode::Section(n) => n.span.as_ref(),
            IRNode::Paragraph(n) => n.span.as_ref(),
            IRNode::Run(n) => n.span.as_ref(),
            IRNode::Hyperlink(n) => n.span.as_ref(),
            IRNode::Table(n) => n.span.as_ref(),
            IRNode::TableRow(n) => n.span.as_ref(),
            IRNode::TableCell(n) => n.span.as_ref(),
            IRNode::Slide(n) => n.span.as_ref(),
            IRNode::Shape(n) => n.span.as_ref(),
            IRNode::Worksheet(n) => n.span.as_ref(),
            IRNode::Cell(n) => n.span.as_ref(),
            IRNode::SharedStringTable(n) => n.span.as_ref(),
            IRNode::SpreadsheetStyles(n) => n.span.as_ref(),
            IRNode::DefinedName(n) => n.span.as_ref(),
            IRNode::ConditionalFormat(n) => n.span.as_ref(),
            IRNode::DataValidation(n) => n.span.as_ref(),
            IRNode::TableDefinition(n) => n.span.as_ref(),
            IRNode::PivotTable(n) => n.span.as_ref(),
            IRNode::PivotCache(n) => n.span.as_ref(),
            IRNode::PivotCacheRecords(n) => n.span.as_ref(),
            IRNode::CalcChain(n) => n.span.as_ref(),
            IRNode::SheetComment(n) => n.span.as_ref(),
            IRNode::SheetMetadata(n) => n.span.as_ref(),
            IRNode::WorkbookProperties(n) => n.span.as_ref(),
            IRNode::MacroProject(n) => n.span.as_ref(),
            IRNode::MacroModule(n) => n.span.as_ref(),
            IRNode::OleObject(n) => n.span.as_ref(),
            IRNode::ExternalReference(n) => n.span.as_ref(),
            IRNode::ActiveXControl(n) => n.span.as_ref(),
            IRNode::Metadata(_) => None,
            IRNode::StyleSet(n) => n.span.as_ref(),
            IRNode::NumberingSet(n) => n.span.as_ref(),
            IRNode::Comment(n) => n.span.as_ref(),
            IRNode::CommentRangeStart(n) => n.span.as_ref(),
            IRNode::CommentRangeEnd(n) => n.span.as_ref(),
            IRNode::CommentReference(n) => n.span.as_ref(),
            IRNode::Footnote(n) => n.span.as_ref(),
            IRNode::Endnote(n) => n.span.as_ref(),
            IRNode::Header(n) => n.span.as_ref(),
            IRNode::Footer(n) => n.span.as_ref(),
            IRNode::WordSettings(n) => n.span.as_ref(),
            IRNode::WebSettings(n) => n.span.as_ref(),
            IRNode::FontTable(n) => n.span.as_ref(),
            IRNode::ContentControl(n) => n.span.as_ref(),
            IRNode::BookmarkStart(n) => n.span.as_ref(),
            IRNode::BookmarkEnd(n) => n.span.as_ref(),
            IRNode::Field(n) => n.span.as_ref(),
            IRNode::Revision(n) => n.span.as_ref(),
            IRNode::CommentExtensionSet(n) => n.span.as_ref(),
            IRNode::CommentIdMap(n) => n.span.as_ref(),
            IRNode::SlideMaster(n) => n.span.as_ref(),
            IRNode::SlideLayout(n) => n.span.as_ref(),
            IRNode::NotesMaster(n) => n.span.as_ref(),
            IRNode::HandoutMaster(n) => n.span.as_ref(),
            IRNode::NotesSlide(n) => n.span.as_ref(),
            IRNode::WorksheetDrawing(n) => n.span.as_ref(),
            IRNode::ChartData(n) => n.span.as_ref(),
            IRNode::PresentationProperties(n) => n.span.as_ref(),
            IRNode::ViewProperties(n) => n.span.as_ref(),
            IRNode::TableStyleSet(n) => n.span.as_ref(),
            IRNode::PptxCommentAuthor(n) => n.span.as_ref(),
            IRNode::PptxComment(n) => n.span.as_ref(),
            IRNode::PresentationTag(n) => n.span.as_ref(),
            IRNode::PresentationInfo(n) => n.span.as_ref(),
            IRNode::PeoplePart(n) => n.span.as_ref(),
            IRNode::SmartArtPart(n) => n.span.as_ref(),
            IRNode::WebExtension(n) => n.span.as_ref(),
            IRNode::WebExtensionTaskpane(n) => n.span.as_ref(),
            IRNode::GlossaryDocument(n) => n.span.as_ref(),
            IRNode::GlossaryEntry(n) => n.span.as_ref(),
            IRNode::VmlDrawing(n) => n.span.as_ref(),
            IRNode::VmlShape(n) => n.span.as_ref(),
            IRNode::DrawingPart(n) => n.span.as_ref(),
            IRNode::ExternalLinkPart(n) => n.span.as_ref(),
            IRNode::ConnectionPart(n) => n.span.as_ref(),
            IRNode::SlicerPart(n) => n.span.as_ref(),
            IRNode::TimelinePart(n) => n.span.as_ref(),
            IRNode::QueryTablePart(n) => n.span.as_ref(),
            IRNode::Diagnostics(n) => n.span.as_ref(),
            IRNode::Theme(n) => n.span.as_ref(),
            IRNode::MediaAsset(n) => n.span.as_ref(),
            IRNode::CustomXmlPart(n) => n.span.as_ref(),
            IRNode::RelationshipGraph(n) => n.span.as_ref(),
            IRNode::DigitalSignature(n) => n.span.as_ref(),
            IRNode::ExtensionPart(n) => n.span.as_ref(),
        }
    }
}
