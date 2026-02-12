//! Core types used throughout the IR.

use serde::{Deserialize, Serialize};
use std::fmt;
use std::str::FromStr;
use std::sync::atomic::{AtomicU64, Ordering};

/// Global counter for generating unique node IDs.
static NODE_ID_COUNTER: AtomicU64 = AtomicU64::new(1);

/// Parse errors for enum string conversions.
#[derive(Debug, Clone)]
pub struct ParseEnumError {
    kind: &'static str,
    input: String,
}

impl ParseEnumError {
    fn new(kind: &'static str, input: &str) -> Self {
        Self {
            kind,
            input: input.to_string(),
        }
    }
}

impl fmt::Display for ParseEnumError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "Unknown {}: {}", self.kind, self.input)
    }
}

impl std::error::Error for ParseEnumError {}

/// Unique identifier for IR nodes.
///
/// NodeIds are stable across serialization and can be used
/// to reference nodes within the IR tree.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
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
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
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
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
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

impl DocumentFormat {
    /// Returns the typical file extension for this format.
    pub fn extension(&self) -> &'static str {
        match self {
            Self::WordProcessing => "docx",
            Self::Spreadsheet => "xlsx",
            Self::Presentation => "pptx",
            Self::OdfText => "odt",
            Self::OdfSpreadsheet => "ods",
            Self::OdfPresentation => "odp",
            Self::Hwp => "hwp",
            Self::Hwpx => "hwpx",
            Self::Rtf => "rtf",
        }
    }

    /// Returns a human-readable name for this format.
    pub fn display_name(&self) -> &'static str {
        match self {
            Self::WordProcessing => "Word Document",
            Self::Spreadsheet => "Excel Spreadsheet",
            Self::Presentation => "PowerPoint Presentation",
            Self::OdfText => "OpenDocument Text",
            Self::OdfSpreadsheet => "OpenDocument Spreadsheet",
            Self::OdfPresentation => "OpenDocument Presentation",
            Self::Hwp => "Hangul Word Processor (HWP)",
            Self::Hwpx => "Hangul Word Processor (HWPX)",
            Self::Rtf => "Rich Text Format (RTF)",
        }
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

/// Node type discriminant for the IR tree.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
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
            "DOCUMENT" => NodeType::Document,
            "SECTION" => NodeType::Section,
            "PARAGRAPH" => NodeType::Paragraph,
            "RUN" => NodeType::Run,
            "TABLE" => NodeType::Table,
            "TABLEROW" | "TABLE_ROW" => NodeType::TableRow,
            "TABLECELL" | "TABLE_CELL" => NodeType::TableCell,
            "SLIDE" => NodeType::Slide,
            "SHAPE" => NodeType::Shape,
            "WORKSHEET" => NodeType::Worksheet,
            "CELL" => NodeType::Cell,
            "MACROPROJECT" | "MACRO_PROJECT" => NodeType::MacroProject,
            "MACROMODULE" | "MACRO_MODULE" => NodeType::MacroModule,
            "OLEOBJECT" | "OLE_OBJECT" => NodeType::OleObject,
            "EXTERNALREFERENCE" | "EXTERNAL_REFERENCE" => NodeType::ExternalReference,
            "ACTIVEXCONTROL" | "ACTIVEX_CONTROL" => NodeType::ActiveXControl,
            _ => return Err(ParseEnumError::new("node type", input)),
        };
        Ok(ty)
    }
}

impl fmt::Display for NodeType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:?}", self)
    }
}
