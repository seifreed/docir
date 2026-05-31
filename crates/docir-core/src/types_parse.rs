use super::{DocumentFormat, NodeType, ParseEnumError};
use std::str::FromStr;

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

#[cfg(test)]
mod tests {
    use super::{NodeType, parse_node_type};

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
