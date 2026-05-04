//! Core types used throughout the IR.

#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};
use std::fmt;
use std::str::FromStr;
use std::sync::atomic::{AtomicU64, Ordering};

/// Global counter for generating unique node IDs.
static NODE_ID_COUNTER: AtomicU64 = AtomicU64::new(1);

/// Parse errors for enum string conversions.
#[derive(Debug, Clone, thiserror::Error)]
#[error("Unknown {kind}: {input}")]
pub struct ParseEnumError {
    kind: &'static str,
    input: String,
}

impl ParseEnumError {
    pub(crate) fn new(kind: &'static str, input: &str) -> Self {
        Self {
            kind,
            input: input.to_string(),
        }
    }
}

/// Unique identifier for IR nodes.
///
/// NodeIds are stable across serialization and can be used
/// to reference nodes within the IR tree.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct NodeId(u64);

impl NodeId {
    /// Creates a new unique NodeId.
    pub fn new() -> Self {
        Self(NODE_ID_COUNTER.fetch_add(1, Ordering::SeqCst))
    }

    /// Creates a NodeId from a raw value (for deserialization).
    pub fn from_raw(value: u64) -> Self {
        Self(value)
    }

    /// Returns the raw u64 value.
    pub fn as_u64(&self) -> u64 {
        self.0
    }
}

impl Default for NodeId {
    fn default() -> Self {
        Self::new()
    }
}

impl fmt::Display for NodeId {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "node_{:08x}", self.0)
    }
}

/// Source location in the original OOXML file.
///
/// Used for diagnostics and tracing parsed elements back to their origin.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceSpan {
    /// Path within the OOXML package (e.g., "word/document.xml").
    pub file_path: String,

    /// Relationship ID if applicable.
    pub relationship_id: Option<String>,

    /// XPath-like path to the element.
    pub xml_path: Option<String>,

    /// Line number in the XML file (if available).
    pub line: Option<u32>,

    /// Column number in the XML file (if available).
    pub column: Option<u32>,
}

impl SourceSpan {
    /// Creates a new SourceSpan with the given file path.
    pub fn new(file_path: impl Into<String>) -> Self {
        Self {
            file_path: file_path.into(),
            relationship_id: None,
            xml_path: None,
            line: None,
            column: None,
        }
    }

    /// Adds a relationship ID to the span.
    pub fn with_relationship(mut self, rel_id: impl Into<String>) -> Self {
        self.relationship_id = Some(rel_id.into());
        self
    }

    /// Adds an XML path to the span.
    pub fn with_xml_path(mut self, path: impl Into<String>) -> Self {
        self.xml_path = Some(path.into());
        self
    }

    /// Adds line and column information.
    pub fn with_position(mut self, line: u32, column: u32) -> Self {
        self.line = Some(line);
        self.column = Some(column);
        self
    }
}

/// Document format type.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DocumentFormat {
    /// Word document (.docx, .docm)
    WordProcessing,
    /// Excel spreadsheet (.xlsx, .xlsm)
    Spreadsheet,
    /// PowerPoint presentation (.pptx, .pptm)
    Presentation,
    /// OpenDocument text (.odt)
    OdfText,
    /// OpenDocument spreadsheet (.ods)
    OdfSpreadsheet,
    /// OpenDocument presentation (.odp)
    OdfPresentation,
    /// Hangul Word Processor legacy format (.hwp)
    Hwp,
    /// Hangul Word Processor XML format (.hwpx)
    Hwpx,
    /// Rich Text Format (.rtf)
    Rtf,
}

struct DocumentFormatDescriptor {
    extension: &'static str,
    display_name: &'static str,
}

impl DocumentFormat {
    fn descriptor(&self) -> DocumentFormatDescriptor {
        match self {
            Self::WordProcessing => DocumentFormatDescriptor {
                extension: "docx",
                display_name: "Word Document",
            },
            Self::Spreadsheet => DocumentFormatDescriptor {
                extension: "xlsx",
                display_name: "Excel Spreadsheet",
            },
            Self::Presentation => DocumentFormatDescriptor {
                extension: "pptx",
                display_name: "PowerPoint Presentation",
            },
            Self::OdfText => DocumentFormatDescriptor {
                extension: "odt",
                display_name: "OpenDocument Text",
            },
            Self::OdfSpreadsheet => DocumentFormatDescriptor {
                extension: "ods",
                display_name: "OpenDocument Spreadsheet",
            },
            Self::OdfPresentation => DocumentFormatDescriptor {
                extension: "odp",
                display_name: "OpenDocument Presentation",
            },
            Self::Hwp => DocumentFormatDescriptor {
                extension: "hwp",
                display_name: "Hangul Word Processor (HWP)",
            },
            Self::Hwpx => DocumentFormatDescriptor {
                extension: "hwpx",
                display_name: "Hangul Word Processor (HWPX)",
            },
            Self::Rtf => DocumentFormatDescriptor {
                extension: "rtf",
                display_name: "Rich Text Format (RTF)",
            },
        }
    }

    /// Returns the typical file extension for this format.
    pub fn extension(&self) -> &'static str {
        self.descriptor().extension
    }

    /// Returns a human-readable name for this format.
    pub fn display_name(&self) -> &'static str {
        self.descriptor().display_name
    }
}

impl FromStr for DocumentFormat {
    type Err = ParseEnumError;

    fn from_str(input: &str) -> Result<Self, Self::Err> {
        let upper = input.trim().to_ascii_uppercase();
        let fmt = match upper.as_str() {
            "DOCX" | "WORD" | "WORDPROCESSING" => DocumentFormat::WordProcessing,
            "XLSX" | "EXCEL" | "SPREADSHEET" => DocumentFormat::Spreadsheet,
            "PPTX" | "PPT" | "POWERPOINT" | "PRESENTATION" => DocumentFormat::Presentation,
            "ODT" | "ODF" | "ODFTEXT" => DocumentFormat::OdfText,
            "ODS" | "ODFSPREADSHEET" => DocumentFormat::OdfSpreadsheet,
            "ODP" | "ODFPRESENTATION" => DocumentFormat::OdfPresentation,
            "HWP" => DocumentFormat::Hwp,
            "HWPX" => DocumentFormat::Hwpx,
            "RTF" => DocumentFormat::Rtf,
            _ => return Err(ParseEnumError::new("document format", input)),
        };
        Ok(fmt)
    }
}

/// Parses a document format from string input.
pub fn parse_document_format(input: &str) -> Result<DocumentFormat, ParseEnumError> {
    DocumentFormat::from_str(input)
}

/// Node type discriminant for the IR tree.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum NodeType {
    // Document root
    Document,

    // Common structural nodes
    Section,
    Paragraph,
    Run,
    Text,
    Hyperlink,

    // Table nodes
    Table,
    TableRow,
    TableCell,

    // Presentation nodes
    Slide,
    Shape,
    TextFrame,

    // Spreadsheet nodes
    Worksheet,
    Cell,
    Formula,
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

    // Security-related nodes
    MacroProject,
    MacroModule,
    OleObject,
    ExternalReference,
    ActiveXControl,

    // Metadata
    Metadata,
    CustomProperty,

    // Media
    Image,
    EmbeddedMedia,

    // DOCX specific
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

    // Shared/package nodes
    Theme,
    MediaAsset,
    CustomXmlPart,
    RelationshipGraph,
    DigitalSignature,
    ExtensionPart,
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
}

impl FromStr for NodeType {
    type Err = ParseEnumError;

    fn from_str(input: &str) -> Result<Self, Self::Err> {
        let upper = input.trim().to_ascii_uppercase();
        parse_node_type_name(&upper).ok_or_else(|| ParseEnumError::new("node type", input))
    }
}

fn parse_node_type_name(input: &str) -> Option<NodeType> {
    parse_common_node_type_name(input)
        .or_else(|| parse_spreadsheet_node_type_name(input))
        .or_else(|| parse_security_node_type_name(input))
        .or_else(|| parse_docx_node_type_name(input))
        .or_else(|| parse_package_node_type_name(input))
}

fn parse_common_node_type_name(input: &str) -> Option<NodeType> {
    match input {
        "DOCUMENT" => Some(NodeType::Document),
        "SECTION" => Some(NodeType::Section),
        "PARAGRAPH" => Some(NodeType::Paragraph),
        "RUN" => Some(NodeType::Run),
        "TEXT" => Some(NodeType::Text),
        "HYPERLINK" => Some(NodeType::Hyperlink),
        "TABLE" => Some(NodeType::Table),
        "TABLEROW" | "TABLE_ROW" => Some(NodeType::TableRow),
        "TABLECELL" | "TABLE_CELL" => Some(NodeType::TableCell),
        "SLIDE" => Some(NodeType::Slide),
        "SHAPE" => Some(NodeType::Shape),
        "TEXTFRAME" | "TEXT_FRAME" => Some(NodeType::TextFrame),
        _ => None,
    }
}

fn parse_spreadsheet_node_type_name(input: &str) -> Option<NodeType> {
    match input {
        "WORKSHEET" => Some(NodeType::Worksheet),
        "CELL" => Some(NodeType::Cell),
        "FORMULA" => Some(NodeType::Formula),
        "SHAREDSTRINGTABLE" | "SHARED_STRING_TABLE" => Some(NodeType::SharedStringTable),
        "SPREADSHEETSTYLES" | "SPREADSHEET_STYLES" => Some(NodeType::SpreadsheetStyles),
        "DEFINEDNAME" | "DEFINED_NAME" => Some(NodeType::DefinedName),
        "CONDITIONALFORMAT" | "CONDITIONAL_FORMAT" => Some(NodeType::ConditionalFormat),
        "DATAVALIDATION" | "DATA_VALIDATION" => Some(NodeType::DataValidation),
        "TABLEDEFINITION" | "TABLE_DEFINITION" => Some(NodeType::TableDefinition),
        "PIVOTTABLE" | "PIVOT_TABLE" => Some(NodeType::PivotTable),
        "PIVOTCACHE" | "PIVOT_CACHE" => Some(NodeType::PivotCache),
        "PIVOTCACHERECORDS" | "PIVOT_CACHE_RECORDS" => Some(NodeType::PivotCacheRecords),
        "CALCCHAIN" | "CALC_CHAIN" => Some(NodeType::CalcChain),
        "SHEETCOMMENT" | "SHEET_COMMENT" => Some(NodeType::SheetComment),
        "SHEETMETADATA" | "SHEET_METADATA" => Some(NodeType::SheetMetadata),
        "WORKBOOKPROPERTIES" | "WORKBOOK_PROPERTIES" => Some(NodeType::WorkbookProperties),
        _ => None,
    }
}

fn parse_security_node_type_name(input: &str) -> Option<NodeType> {
    match input {
        "MACROPROJECT" | "MACRO_PROJECT" => Some(NodeType::MacroProject),
        "MACROMODULE" | "MACRO_MODULE" => Some(NodeType::MacroModule),
        "OLEOBJECT" | "OLE_OBJECT" => Some(NodeType::OleObject),
        "EXTERNALREFERENCE" | "EXTERNAL_REFERENCE" => Some(NodeType::ExternalReference),
        "ACTIVEXCONTROL" | "ACTIVEX_CONTROL" => Some(NodeType::ActiveXControl),
        "METADATA" => Some(NodeType::Metadata),
        "CUSTOMPROPERTY" | "CUSTOM_PROPERTY" => Some(NodeType::CustomProperty),
        "IMAGE" => Some(NodeType::Image),
        "EMBEDDEDMEDIA" | "EMBEDDED_MEDIA" => Some(NodeType::EmbeddedMedia),
        _ => None,
    }
}

fn parse_docx_node_type_name(input: &str) -> Option<NodeType> {
    match input {
        "STYLESET" | "STYLE_SET" => Some(NodeType::StyleSet),
        "NUMBERINGSET" | "NUMBERING_SET" => Some(NodeType::NumberingSet),
        "COMMENT" => Some(NodeType::Comment),
        "COMMENTRANGESTART" | "COMMENT_RANGE_START" => Some(NodeType::CommentRangeStart),
        "COMMENTRANGEEND" | "COMMENT_RANGE_END" => Some(NodeType::CommentRangeEnd),
        "COMMENTREFERENCE" | "COMMENT_REFERENCE" => Some(NodeType::CommentReference),
        "FOOTNOTE" => Some(NodeType::Footnote),
        "ENDNOTE" => Some(NodeType::Endnote),
        "HEADER" => Some(NodeType::Header),
        "FOOTER" => Some(NodeType::Footer),
        "WORDSETTINGS" | "WORD_SETTINGS" => Some(NodeType::WordSettings),
        "WEBSETTINGS" | "WEB_SETTINGS" => Some(NodeType::WebSettings),
        "FONTTABLE" | "FONT_TABLE" => Some(NodeType::FontTable),
        "CONTENTCONTROL" | "CONTENT_CONTROL" => Some(NodeType::ContentControl),
        "BOOKMARKSTART" | "BOOKMARK_START" => Some(NodeType::BookmarkStart),
        "BOOKMARKEND" | "BOOKMARK_END" => Some(NodeType::BookmarkEnd),
        "FIELD" => Some(NodeType::Field),
        "REVISION" => Some(NodeType::Revision),
        "COMMENTEXTENSIONSET" | "COMMENT_EXTENSION_SET" => Some(NodeType::CommentExtensionSet),
        "COMMENTIDMAP" | "COMMENT_ID_MAP" => Some(NodeType::CommentIdMap),
        "GLOSSARYDOCUMENT" | "GLOSSARY_DOCUMENT" => Some(NodeType::GlossaryDocument),
        "GLOSSARYENTRY" | "GLOSSARY_ENTRY" => Some(NodeType::GlossaryEntry),
        _ => None,
    }
}

fn parse_package_node_type_name(input: &str) -> Option<NodeType> {
    match input {
        "THEME" => Some(NodeType::Theme),
        "MEDIAASSET" | "MEDIA_ASSET" => Some(NodeType::MediaAsset),
        "CUSTOMXMLPART" | "CUSTOM_XML_PART" => Some(NodeType::CustomXmlPart),
        "RELATIONSHIPGRAPH" | "RELATIONSHIP_GRAPH" => Some(NodeType::RelationshipGraph),
        "DIGITALSIGNATURE" | "DIGITAL_SIGNATURE" => Some(NodeType::DigitalSignature),
        "EXTENSIONPART" | "EXTENSION_PART" => Some(NodeType::ExtensionPart),
        "SLIDEMASTER" | "SLIDE_MASTER" => Some(NodeType::SlideMaster),
        "SLIDELAYOUT" | "SLIDE_LAYOUT" => Some(NodeType::SlideLayout),
        "NOTESMASTER" | "NOTES_MASTER" => Some(NodeType::NotesMaster),
        "HANDOUTMASTER" | "HANDOUT_MASTER" => Some(NodeType::HandoutMaster),
        "NOTESSLIDE" | "NOTES_SLIDE" => Some(NodeType::NotesSlide),
        "WORKSHEETDRAWING" | "WORKSHEET_DRAWING" => Some(NodeType::WorksheetDrawing),
        "CHARTDATA" | "CHART_DATA" => Some(NodeType::ChartData),
        "PRESENTATIONPROPERTIES" | "PRESENTATION_PROPERTIES" => {
            Some(NodeType::PresentationProperties)
        }
        "VIEWPROPERTIES" | "VIEW_PROPERTIES" => Some(NodeType::ViewProperties),
        "TABLESTYLESET" | "TABLE_STYLE_SET" => Some(NodeType::TableStyleSet),
        "PPTXCOMMENTAUTHOR" | "PPTX_COMMENT_AUTHOR" => Some(NodeType::PptxCommentAuthor),
        "PPTXCOMMENT" | "PPTX_COMMENT" => Some(NodeType::PptxComment),
        "PRESENTATIONTAG" | "PRESENTATION_TAG" => Some(NodeType::PresentationTag),
        "PRESENTATIONINFO" | "PRESENTATION_INFO" => Some(NodeType::PresentationInfo),
        "PEOPLEPART" | "PEOPLE_PART" => Some(NodeType::PeoplePart),
        "SMARTARTPART" | "SMART_ART_PART" => Some(NodeType::SmartArtPart),
        "WEBEXTENSION" | "WEB_EXTENSION" => Some(NodeType::WebExtension),
        "WEBEXTENSIONTASKPANE" | "WEB_EXTENSION_TASKPANE" => Some(NodeType::WebExtensionTaskpane),
        "VMLDRAWING" | "VML_DRAWING" => Some(NodeType::VmlDrawing),
        "VMLSHAPE" | "VML_SHAPE" => Some(NodeType::VmlShape),
        "DRAWINGPART" | "DRAWING_PART" => Some(NodeType::DrawingPart),
        "EXTERNALLINKPART" | "EXTERNAL_LINK_PART" => Some(NodeType::ExternalLinkPart),
        "CONNECTIONPART" | "CONNECTION_PART" => Some(NodeType::ConnectionPart),
        "SLICERPART" | "SLICER_PART" => Some(NodeType::SlicerPart),
        "TIMELINEPART" | "TIMELINE_PART" => Some(NodeType::TimelinePart),
        "QUERYTABLEPART" | "QUERY_TABLE_PART" => Some(NodeType::QueryTablePart),
        "DIAGNOSTICS" => Some(NodeType::Diagnostics),
        _ => None,
    }
}

/// Parses a node type from string input.
pub fn parse_node_type(input: &str) -> Result<NodeType, ParseEnumError> {
    NodeType::from_str(input)
}

impl fmt::Display for NodeType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:?}", self)
    }
}

#[cfg(test)]
mod tests {
    use super::{parse_node_type, NodeType};

    #[test]
    fn parse_node_type_accepts_representative_aliases() {
        let cases = [
            ("text", NodeType::Text),
            ("table_row", NodeType::TableRow),
            ("pivot_cache_records", NodeType::PivotCacheRecords),
            ("macro_project", NodeType::MacroProject),
            ("comment_range_start", NodeType::CommentRangeStart),
            ("presentation_properties", NodeType::PresentationProperties),
            ("query_table_part", NodeType::QueryTablePart),
        ];

        for (input, expected) in cases {
            assert_eq!(parse_node_type(input).expect("node type alias"), expected);
        }
    }
}
