//! [Content_Types].xml parser.

use crate::error::ParseError;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;

/// Content types registry from [Content_Types].xml.
#[derive(Debug, Clone, Default)]
pub struct ContentTypes {
    /// Default content types by extension.
    pub defaults: HashMap<String, String>,
    /// Override content types by part name.
    pub overrides: HashMap<String, String>,
}

impl ContentTypes {
    /// Parses [Content_Types].xml content.
    pub fn parse(xml: &str) -> Result<Self, ParseError> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut content_types = ContentTypes::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    match e.name().as_ref() {
                        b"Default" => {
                            let mut extension = None;
                            let mut content_type = None;

                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"Extension" => {
                                        extension =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                    b"ContentType" => {
                                        content_type =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                    _ => {}
                                }
                            }

                            if let (Some(ext), Some(ct)) = (extension, content_type) {
                                content_types.defaults.insert(ext, ct);
                            }
                        }
                        b"Override" => {
                            let mut part_name = None;
                            let mut content_type = None;

                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"PartName" => {
                                        part_name =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                    b"ContentType" => {
                                        content_type =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                    _ => {}
                                }
                            }

                            if let (Some(pn), Some(ct)) = (part_name, content_type) {
                                // Remove leading slash for consistency
                                let normalized = pn.strip_prefix('/').unwrap_or(&pn).to_string();
                                content_types.overrides.insert(normalized, ct);
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: "[Content_Types].xml".to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(content_types)
    }

    /// Gets the content type for a given part.
    pub fn get_content_type(&self, part_name: &str) -> Option<&str> {
        // Remove leading slash for lookup
        let normalized = part_name.strip_prefix('/').unwrap_or(part_name);

        // Check overrides first
        if let Some(ct) = self.overrides.get(normalized) {
            return Some(ct);
        }

        // Fall back to defaults by extension
        if let Some(ext) = normalized.rsplit('.').next() {
            let ext = ext.to_ascii_lowercase();
            if let Some(ct) = self.defaults.get(&ext) {
                return Some(ct);
            }
        }

        None
    }

    /// Checks if this is a macro-enabled document.
    pub fn is_macro_enabled(&self) -> bool {
        self.overrides
            .values()
            .any(|ct| ct.contains("macroEnabled"))
    }

    /// Detects the document format from content types.
    pub fn detect_format(&self) -> Option<docir_core::DocumentFormat> {
        for ct in self.overrides.values() {
            if ct.contains("wordprocessingml") {
                return Some(docir_core::DocumentFormat::WordProcessing);
            }
            if ct.contains("spreadsheetml") {
                return Some(docir_core::DocumentFormat::Spreadsheet);
            }
            if ct.contains("sheet.binary") {
                return Some(docir_core::DocumentFormat::Spreadsheet);
            }
            if ct.contains("presentationml") {
                return Some(docir_core::DocumentFormat::Presentation);
            }
        }
        None
    }

    /// Returns true if the part is treated as a legacy/extension part.
    pub fn is_extension_part(&self, part_name: &str) -> bool {
        self.get_content_type(part_name)
            .map(|ct| ct.contains("extension") && !ct.contains("webextension"))
            .unwrap_or(false)
    }
}

/// Known OOXML content types.
pub mod content_type {
    pub const CORE_PROPERTIES: &str = "application/vnd.openxmlformats-package.core-properties+xml";
    pub const EXTENDED_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    pub const CUSTOM_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.custom-properties+xml";
    pub const CUSTOM_XML: &str = "application/xml";
    pub const CUSTOM_XML_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
    pub const DIGITAL_SIGNATURE_XML: &str =
        "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";
    pub const DIGITAL_SIGNATURE_ORIGIN: &str =
        "application/vnd.openxmlformats-package.digital-signature-origin";
    pub const WORD_DOCUMENT: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    pub const WORD_DOCUMENT_MACRO: &str = "application/vnd.ms-word.document.macroEnabled.main+xml";
    pub const WORD_STYLES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    pub const WORD_NUMBERING: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
    pub const WORD_GLOSSARY: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";
    pub const WORD_SETTINGS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    pub const WORD_WEB_SETTINGS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml";
    pub const WORD_FONT_TABLE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
    pub const WORD_COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    pub const WORD_COMMENTS_EXTENDED: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
    pub const WORD_COMMENTS_IDS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml";
    pub const WORD_FOOTNOTES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
    pub const WORD_ENDNOTES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
    pub const WORD_HEADER: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    pub const WORD_FOOTER: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
    pub const WORD_PEOPLE: &str = "application/vnd.openxmlformats-officedocument.people+xml";
    pub const WORD_THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
    pub const WORD_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
    pub const WORD_DIAGRAM_DATA: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
    pub const WORD_DIAGRAM_LAYOUT: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml";
    pub const WORD_DIAGRAM_COLORS: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml";
    pub const WORD_DIAGRAM_STYLE: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml";
    pub const WORD_CHART: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    pub const WORD_VML_DRAWING: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";
    pub const WORD_VBA_DATA: &str = "application/vnd.ms-office.vbaData+xml";
    pub const WORD_ACTIVEX_XML: &str = "application/vnd.ms-office.activeX+xml";
    pub const WORD_ACTIVEX_BIN: &str = "application/vnd.ms-office.activeX";
    pub const WORD_WEB_EXTENSION: &str = "application/vnd.ms-office.webextension+xml";
    pub const WORD_WEB_EXTENSION_TASKPANES: &str =
        "application/vnd.ms-office.webextensiontaskpanes+xml";

    pub const EXCEL_WORKBOOK: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    pub const EXCEL_WORKBOOK_MACRO: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
    pub const EXCEL_WORKBOOK_BIN: &str = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
    pub const EXCEL_WORKSHEET: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    pub const EXCEL_CHARTSHEET: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
    pub const EXCEL_DIALOGSHEET: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml";
    pub const EXCEL_MACROSHEET: &str = "application/vnd.ms-excel.macrosheet+xml";
    pub const EXCEL_SHARED_STRINGS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
    pub const EXCEL_STYLES: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    pub const EXCEL_CALC_CHAIN: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
    pub const EXCEL_TABLE: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
    pub const EXCEL_PIVOT_TABLE: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
    pub const EXCEL_PIVOT_CACHE_DEF: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
    pub const EXCEL_PIVOT_CACHE_RECORDS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
    pub const EXCEL_QUERY_TABLE: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
    pub const EXCEL_CONNECTIONS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
    pub const EXCEL_EXTERNAL_LINK: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
    pub const EXCEL_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
    pub const EXCEL_CHART: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    pub const EXCEL_SHEET_METADATA: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
    pub const EXCEL_PERSON: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.person+xml";
    pub const EXCEL_SLICER: &str = "application/vnd.ms-excel.slicer+xml";
    pub const EXCEL_TIMELINE: &str = "application/vnd.ms-excel.timeline+xml";
    pub const EXCEL_COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
    pub const EXCEL_THREADED_COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.threadedComments+xml";
    pub const EXCEL_ACTIVEX_XML: &str = "application/vnd.ms-office.activeX+xml";
    pub const EXCEL_ACTIVEX_BIN: &str = "application/vnd.ms-office.activeX";

    pub const PPTX_PRESENTATION: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    pub const PPTX_PRESENTATION_MACRO: &str =
        "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml";
    pub const PPTX_SLIDE: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
    pub const PPTX_PRES_PROPS: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";
    pub const PPTX_VIEW_PROPS: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml";
    pub const PPTX_TABLE_STYLES: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";
    pub const PPTX_SLIDE_MASTER: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
    pub const PPTX_SLIDE_LAYOUT: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
    pub const PPTX_NOTES_MASTER: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
    pub const PPTX_HANDOUT_MASTER: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml";
    pub const PPTX_NOTES_SLIDE: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
    pub const PPTX_DIAGRAM_DATA: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
    pub const PPTX_DIAGRAM_LAYOUT: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml";
    pub const PPTX_DIAGRAM_COLORS: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml";
    pub const PPTX_DIAGRAM_STYLE: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml";
    pub const PPTX_CHART: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    pub const PPTX_COMMENT_AUTHORS: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
    pub const PPTX_COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
    pub const PPTX_PEOPLE: &str = "application/vnd.openxmlformats-officedocument.people+xml";
    pub const PPTX_TAGS: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";
    pub const PPTX_ACTIVEX_XML: &str = "application/vnd.ms-office.activeX+xml";
    pub const PPTX_ACTIVEX_BIN: &str = "application/vnd.ms-office.activeX";

    pub const VBA_PROJECT: &str = "application/vnd.ms-office.vbaProject";
    pub const OLE_OBJECT: &str = "application/vnd.openxmlformats-officedocument.oleObject";
    pub const RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";
}
