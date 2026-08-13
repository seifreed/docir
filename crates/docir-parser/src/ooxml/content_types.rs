//! [Content_Types].xml parser.

use crate::error::ParseError;
use crate::xml_utils::{
    local_name, read_event, reader_from_str, track_xml_document_event, try_attr_value,
};
use quick_xml::events::Event;
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
        let mut reader = reader_from_str(xml);

        let mut content_types = ContentTypes::default();
        let mut buf = Vec::new();
        let mut depth = 0usize;
        let mut root_closed = false;

        loop {
            let event = read_event(&mut reader, &mut buf, "[Content_Types].xml")?;
            if track_xml_document_event(
                &event,
                &mut depth,
                &mut root_closed,
                "[Content_Types].xml",
            )? {
                break;
            }
            match event {
                Event::Empty(e) | Event::Start(e) => match local_name(e.name().as_ref()) {
                    b"Default" => {
                        let (ext, ct) = parse_default_entry(&e)?;
                        content_types.defaults.insert(ext.to_ascii_lowercase(), ct);
                    }
                    b"Override" => {
                        let (pn, ct) = parse_override_entry(&e)?;
                        let normalized = normalize_part_name(&pn);
                        content_types.overrides.insert(normalized, ct);
                    }
                    _ => {}
                },
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(content_types)
    }

    /// Gets the content type for a given part.
    pub fn get_content_type(&self, part_name: &str) -> Option<&str> {
        // Remove leading slash for lookup
        let normalized = if let Some(without_prefix) = part_name.strip_prefix('/') {
            without_prefix
        } else {
            part_name
        }
        .to_ascii_lowercase();

        // Check overrides first
        if let Some(ct) = self.overrides.get(&normalized) {
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
        self.main_content_type().is_some_and(|ct| {
            matches!(
                ct,
                content_type::WORD_DOCUMENT_MACRO
                    | content_type::EXCEL_WORKBOOK_MACRO
                    | content_type::PPTX_PRESENTATION_MACRO
            )
        })
    }

    /// Detects the document format from content types.
    pub fn detect_format(&self) -> Option<docir_core::DocumentFormat> {
        self.main_content_type()
            .and_then(format_for_main_content_type)
    }

    /// Checks if the package has an XLSB workbook as its main part.
    pub fn is_binary_workbook(&self) -> bool {
        self.get_content_type("xl/workbook.bin")
            .is_some_and(|ct| ct == content_type::EXCEL_WORKBOOK_BIN)
    }

    /// Returns true if the part is treated as a legacy/extension part.
    pub fn is_extension_part(&self, part_name: &str) -> bool {
        if let Some(content_type) = self.get_content_type(part_name) {
            return content_type.contains("extension") && !content_type.contains("webextension");
        }
        false
    }

    fn main_content_type(&self) -> Option<&str> {
        [
            "word/document.xml",
            "xl/workbook.xml",
            "xl/workbook.bin",
            "ppt/presentation.xml",
        ]
        .into_iter()
        .find_map(|part_name| {
            self.get_content_type(part_name)
                .filter(|content_type| format_for_main_content_type(content_type).is_some())
        })
    }
}

fn format_for_main_content_type(content_type: &str) -> Option<docir_core::DocumentFormat> {
    match content_type {
        content_type::WORD_DOCUMENT | content_type::WORD_DOCUMENT_MACRO => {
            Some(docir_core::DocumentFormat::WordProcessing)
        }
        content_type::EXCEL_WORKBOOK
        | content_type::EXCEL_WORKBOOK_MACRO
        | content_type::EXCEL_WORKBOOK_BIN => Some(docir_core::DocumentFormat::Spreadsheet),
        content_type::PPTX_PRESENTATION | content_type::PPTX_PRESENTATION_MACRO => {
            Some(docir_core::DocumentFormat::Presentation)
        }
        _ => None,
    }
}

fn parse_default_entry(
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<(String, String), ParseError> {
    let extension = required_content_type_attr(element, b"Extension")?;
    let content_type = required_content_type_attr(element, b"ContentType")?;
    Ok((extension, content_type))
}

fn parse_override_entry(
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<(String, String), ParseError> {
    let part_name = required_content_type_attr(element, b"PartName")?;
    let content_type = required_content_type_attr(element, b"ContentType")?;
    Ok((part_name, content_type))
}

fn required_content_type_attr(
    element: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
) -> Result<String, ParseError> {
    try_attr_value(element, name, "[Content_Types].xml")?.ok_or_else(|| {
        ParseError::InvalidStructure(format!(
            "[Content_Types].xml entry is missing {}",
            String::from_utf8_lossy(name)
        ))
    })
}

fn normalize_part_name(part_name: &str) -> String {
    part_name
        .strip_prefix('/')
        .unwrap_or(part_name)
        .to_ascii_lowercase()
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

#[cfg(test)]
mod tests {
    use super::{ContentTypes, content_type};
    use crate::error::ParseError;
    use docir_core::DocumentFormat;

    #[test]
    fn parse_accepts_prefixed_content_type_entries() {
        let xml = r#"
        <ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
          <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        </ct:Types>
        "#;

        let types = ContentTypes::parse(xml).expect("content types");

        assert_eq!(
            types.get_content_type("xl/workbook.xml"),
            Some(content_type::EXCEL_WORKBOOK)
        );
        assert_eq!(
            types.get_content_type("xl/_rels/workbook.xml.rels"),
            Some(content_type::RELATIONSHIPS)
        );
        assert_eq!(types.detect_format(), Some(DocumentFormat::Spreadsheet));
    }

    #[test]
    fn detect_format_uses_main_part_default_for_xlsb() {
        let xml = format!(
            r#"
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="bin" ContentType="{}"/>
            </Types>"#,
            content_type::EXCEL_WORKBOOK_BIN
        );

        let types = ContentTypes::parse(&xml).expect("content types");

        assert_eq!(types.detect_format(), Some(DocumentFormat::Spreadsheet));
        assert!(types.is_binary_workbook());
    }

    #[test]
    fn content_types_match_part_names_and_extensions_case_insensitively() {
        let xml = format!(
            r#"
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="BIN" ContentType="{}"/>
              <Override PartName="/XL/WORKBOOK.XML" ContentType="{}"/>
            </Types>"#,
            content_type::EXCEL_WORKBOOK_BIN,
            content_type::EXCEL_WORKBOOK
        );

        let types = ContentTypes::parse(&xml).expect("content types");

        assert_eq!(
            types.get_content_type("xl/workbook.bin"),
            Some(content_type::EXCEL_WORKBOOK)
        );
        assert_eq!(
            types.get_content_type("XL/WORKBOOK.BIN"),
            Some(content_type::EXCEL_WORKBOOK_BIN)
        );
    }

    #[test]
    fn detect_format_ignores_embedded_ooxml_content_types() {
        let xml = format!(
            r#"
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Override PartName="/word/document.xml" ContentType="{}"/>
              <Override PartName="/word/embeddings/workbook.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>
              <Override PartName="/word/embeddings/workbook.xlsm" ContentType="{}"/>
              <Override PartName="/word/embeddings/workbook.xlsb" ContentType="{}"/>
            </Types>"#,
            content_type::WORD_DOCUMENT,
            content_type::EXCEL_WORKBOOK_MACRO,
            content_type::EXCEL_WORKBOOK_BIN
        );

        let types = ContentTypes::parse(&xml).expect("content types");

        assert_eq!(types.detect_format(), Some(DocumentFormat::WordProcessing));
        assert!(!types.is_macro_enabled());
    }

    #[test]
    fn parse_reports_malformed_attributes() {
        let xml = r#"
        <ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
          <ct:Default Extension="rels" Extension="dup" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        </ct:Types>
        "#;

        let err = ContentTypes::parse(xml).expect_err("malformed attributes must fail");

        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "[Content_Types].xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_rejects_incomplete_default_and_override_entries() {
        for xml in [
            r#"
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default ContentType="application/xml"/>
            </Types>
            "#,
            r#"
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Override PartName="/word/document.xml"/>
            </Types>
            "#,
        ] {
            assert!(ContentTypes::parse(xml).is_err());
        }
    }

    #[test]
    fn parse_rejects_truncated_content_types_document() {
        let xml = r#"
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        "#;

        let err = ContentTypes::parse(xml).expect_err("truncated content types must fail");
        assert!(
            matches!(err, ParseError::Xml { file, message } if file == "[Content_Types].xml" && message.contains("root is closed"))
        );
    }
}
