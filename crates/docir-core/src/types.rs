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
        let ty = match upper.as_str() {
            // Structural nodes
            "DOCUMENT" => NodeType::Document,
            "SECTION" => NodeType::Section,
            "PARAGRAPH" => NodeType::Paragraph,
            "RUN" => NodeType::Run,
            "HYPERLINK" => NodeType::Hyperlink,
            // Table nodes
            "TABLE" => NodeType::Table,
            "TABLEROW" | "TABLE_ROW" => NodeType::TableRow,
            "TABLECELL" | "TABLE_CELL" => NodeType::TableCell,
            // Presentation nodes
            "SLIDE" => NodeType::Slide,
            "SHAPE" => NodeType::Shape,
            "TEXTFRAME" | "TEXT_FRAME" => NodeType::TextFrame,
            // Spreadsheet nodes
            "WORKSHEET" => NodeType::Worksheet,
            "CELL" => NodeType::Cell,
            "FORMULA" => NodeType::Formula,
            "SHAREDSTRINGTABLE" | "SHARED_STRING_TABLE" => NodeType::SharedStringTable,
            "SPREADSHEETSTYLES" | "SPREADSHEET_STYLES" => NodeType::SpreadsheetStyles,
            "DEFINEDNAME" | "DEFINED_NAME" => NodeType::DefinedName,
            "CONDITIONALFORMAT" | "CONDITIONAL_FORMAT" => NodeType::ConditionalFormat,
            "DATAVALIDATION" | "DATA_VALIDATION" => NodeType::DataValidation,
            "TABLEDEFINITION" | "TABLE_DEFINITION" => NodeType::TableDefinition,
            "PIVOTTABLE" | "PIVOT_TABLE" => NodeType::PivotTable,
            "PIVOTCACHE" | "PIVOT_CACHE" => NodeType::PivotCache,
            "PIVOTCACHERECORDS" | "PIVOT_CACHE_RECORDS" => NodeType::PivotCacheRecords,
            "CALCCHAIN" | "CALC_CHAIN" => NodeType::CalcChain,
            "SHEETCOMMENT" | "SHEET_COMMENT" => NodeType::SheetComment,
            "SHEETMETADATA" | "SHEET_METADATA" => NodeType::SheetMetadata,
            "WORKBOOKPROPERTIES" | "WORKBOOK_PROPERTIES" => NodeType::WorkbookProperties,
            // Security-related nodes
            "MACROPROJECT" | "MACRO_PROJECT" => NodeType::MacroProject,
            "MACROMODULE" | "MACRO_MODULE" => NodeType::MacroModule,
            "OLEOBJECT" | "OLE_OBJECT" => NodeType::OleObject,
            "EXTERNALREFERENCE" | "EXTERNAL_REFERENCE" => NodeType::ExternalReference,
            "ACTIVEXCONTROL" | "ACTIVEX_CONTROL" => NodeType::ActiveXControl,
            // Metadata nodes
            "METADATA" => NodeType::Metadata,
            "CUSTOMPROPERTY" | "CUSTOM_PROPERTY" => NodeType::CustomProperty,
            // Media nodes
            "IMAGE" => NodeType::Image,
            "EMBEDDEDMEDIA" | "EMBEDDED_MEDIA" => NodeType::EmbeddedMedia,
            // DOCX specific
            "STYLESET" | "STYLE_SET" => NodeType::StyleSet,
            "NUMBERINGSET" | "NUMBERING_SET" => NodeType::NumberingSet,
            "COMMENT" => NodeType::Comment,
            "COMMENTRANGESTART" | "COMMENT_RANGE_START" => NodeType::CommentRangeStart,
            "COMMENTRANGEEND" | "COMMENT_RANGE_END" => NodeType::CommentRangeEnd,
            "COMMENTREFERENCE" | "COMMENT_REFERENCE" => NodeType::CommentReference,
            "FOOTNOTE" => NodeType::Footnote,
            "ENDNOTE" => NodeType::Endnote,
            "HEADER" => NodeType::Header,
            "FOOTER" => NodeType::Footer,
            // Shared/package nodes
            "THEME" => NodeType::Theme,
            "MEDIAASSET" | "MEDIA_ASSET" => NodeType::MediaAsset,
            "CUSTOMXMLPART" | "CUSTOM_XML_PART" => NodeType::CustomXmlPart,
            "RELATIONSHIPGRAPH" | "RELATIONSHIP_GRAPH" => NodeType::RelationshipGraph,
            "DIGITALSIGNATURE" | "DIGITAL_SIGNATURE" => NodeType::DigitalSignature,
            "EXTENSIONPART" | "EXTENSION_PART" => NodeType::ExtensionPart,
            "WORDSETTINGS" | "WORD_SETTINGS" => NodeType::WordSettings,
            "WEBSETTINGS" | "WEB_SETTINGS" => NodeType::WebSettings,
            "FONTTABLE" | "FONT_TABLE" => NodeType::FontTable,
            "CONTENTCONTROL" | "CONTENT_CONTROL" => NodeType::ContentControl,
            "BOOKMARKSTART" | "BOOKMARK_START" => NodeType::BookmarkStart,
            "BOOKMARKEND" | "BOOKMARK_END" => NodeType::BookmarkEnd,
            "FIELD" => NodeType::Field,
            "REVISION" => NodeType::Revision,
            "COMMENTEXTENSIONSET" | "COMMENT_EXTENSION_SET" => NodeType::CommentExtensionSet,
            "COMMENTIDMAP" | "COMMENT_ID_MAP" => NodeType::CommentIdMap,
            "SLIDEMASTER" | "SLIDE_MASTER" => NodeType::SlideMaster,
            "SLIDELAYOUT" | "SLIDE_LAYOUT" => NodeType::SlideLayout,
            "NOTESMASTER" | "NOTES_MASTER" => NodeType::NotesMaster,
            "HANDOUTMASTER" | "HANDOUT_MASTER" => NodeType::HandoutMaster,
            "NOTESSLIDE" | "NOTES_SLIDE" => NodeType::NotesSlide,
            "WORKSHEETDRAWING" | "WORKSHEET_DRAWING" => NodeType::WorksheetDrawing,
            "CHARTDATA" | "CHART_DATA" => NodeType::ChartData,
            "PRESENTATIONPROPERTIES" | "PRESENTATION_PROPERTIES" => {
                NodeType::PresentationProperties
            }
            "VIEWPROPERTIES" | "VIEW_PROPERTIES" => NodeType::ViewProperties,
            "TABLESTYLESET" | "TABLE_STYLE_SET" => NodeType::TableStyleSet,
            "PPTXCOMMENTAUTHOR" | "PPTX_COMMENT_AUTHOR" => NodeType::PptxCommentAuthor,
            "PPTXCOMMENT" | "PPTX_COMMENT" => NodeType::PptxComment,
            "PRESENTATIONTAG" | "PRESENTATION_TAG" => NodeType::PresentationTag,
            "PRESENTATIONINFO" | "PRESENTATION_INFO" => NodeType::PresentationInfo,
            "PEOPLEPART" | "PEOPLE_PART" => NodeType::PeoplePart,
            "SMARTARTPART" | "SMART_ART_PART" => NodeType::SmartArtPart,
            "WEBEXTENSION" | "WEB_EXTENSION" => NodeType::WebExtension,
            "WEBEXTENSIONTASKPANE" | "WEB_EXTENSION_TASKPANE" => NodeType::WebExtensionTaskpane,
            "GLOSSARYDOCUMENT" | "GLOSSARY_DOCUMENT" => NodeType::GlossaryDocument,
            "GLOSSARYENTRY" | "GLOSSARY_ENTRY" => NodeType::GlossaryEntry,
            "VMLDRAWING" | "VML_DRAWING" => NodeType::VmlDrawing,
            "VMLSHAPE" | "VML_SHAPE" => NodeType::VmlShape,
            "DRAWINGPART" | "DRAWING_PART" => NodeType::DrawingPart,
            "EXTERNALLINKPART" | "EXTERNAL_LINK_PART" => NodeType::ExternalLinkPart,
            "CONNECTIONPART" | "CONNECTION_PART" => NodeType::ConnectionPart,
            "SLICERPART" | "SLICER_PART" => NodeType::SlicerPart,
            "TIMELINEPART" | "TIMELINE_PART" => NodeType::TimelinePart,
            "QUERYTABLEPART" | "QUERY_TABLE_PART" => NodeType::QueryTablePart,
            "DIAGNOSTICS" => NodeType::Diagnostics,
            _ => return Err(ParseEnumError::new("node type", input)),
        };
        Ok(ty)
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
