//! OOXML Part Registry for coverage tracking.

use crate::ooxml::content_types::content_type;
use crate::registry_utils::matches_pattern;
use docir_core::types::DocumentFormat;

#[derive(Debug, Clone)]
pub struct PartSpec {
    pub format: DocumentFormat,
    pub pattern: &'static str,
    pub content_type: Option<&'static str>,
    pub expected_parser: &'static str,
}

impl PartSpec {
    pub fn matches(&self, path: &str, content_type_value: Option<&str>) -> bool {
        if let Some(ct) = self.content_type {
            if let Some(actual) = content_type_value {
                if actual != ct {
                    return false;
                }
            }
        }
        matches_pattern(path, self.pattern)
    }
}

#[derive(Debug, Clone, Copy)]
struct PartSpecEntry {
    pattern: &'static str,
    content_type: Option<&'static str>,
    expected_parser: &'static str,
}

fn build_registry(format: DocumentFormat, entries: &[PartSpecEntry]) -> Vec<PartSpec> {
    entries
        .iter()
        .map(|entry| PartSpec {
            format,
            pattern: entry.pattern,
            content_type: entry.content_type,
            expected_parser: entry.expected_parser,
        })
        .collect()
}

const WORD_PARTS: &[PartSpecEntry] = &[
    PartSpecEntry {
        pattern: "_rels/.rels",
        content_type: Some(content_type::RELATIONSHIPS),
        expected_parser: "Relationships::parse",
    },
    PartSpecEntry {
        pattern: "word/_rels/*.rels",
        content_type: Some(content_type::RELATIONSHIPS),
        expected_parser: "Relationships::parse",
    },
    PartSpecEntry {
        pattern: "docProps/core.xml",
        content_type: Some(content_type::CORE_PROPERTIES),
        expected_parser: "parse_metadata(core)",
    },
    PartSpecEntry {
        pattern: "docProps/app.xml",
        content_type: Some(content_type::EXTENDED_PROPERTIES),
        expected_parser: "parse_metadata(app)",
    },
    PartSpecEntry {
        pattern: "docProps/custom.xml",
        content_type: Some(content_type::CUSTOM_PROPERTIES),
        expected_parser: "parse_custom_properties",
    },
    PartSpecEntry {
        pattern: "customXml/*.xml",
        content_type: Some(content_type::CUSTOM_XML),
        expected_parser: "parse_custom_xml_part",
    },
    PartSpecEntry {
        pattern: "customXml/itemProps*.xml",
        content_type: Some(content_type::CUSTOM_XML_PROPERTIES),
        expected_parser: "parse_custom_xml_part",
    },
    PartSpecEntry {
        pattern: "_xmlsignatures/*.xml",
        content_type: Some(content_type::DIGITAL_SIGNATURE_XML),
        expected_parser: "parse_shared_parts(signature)",
    },
    PartSpecEntry {
        pattern: "_xmlsignatures/origin.sigs",
        content_type: Some(content_type::DIGITAL_SIGNATURE_ORIGIN),
        expected_parser: "parse_shared_parts(signature)",
    },
    PartSpecEntry {
        pattern: "word/document.xml",
        content_type: Some(content_type::WORD_DOCUMENT),
        expected_parser: "DocxParser::parse_document",
    },
    PartSpecEntry {
        pattern: "word/document.xml",
        content_type: Some(content_type::WORD_DOCUMENT_MACRO),
        expected_parser: "DocxParser::parse_document",
    },
    PartSpecEntry {
        pattern: "word/glossary/document.xml",
        content_type: Some(content_type::WORD_GLOSSARY),
        expected_parser: "DocxParser::parse_glossary_document",
    },
    PartSpecEntry {
        pattern: "word/styles.xml",
        content_type: Some(content_type::WORD_STYLES),
        expected_parser: "DocxParser::parse_styles",
    },
    PartSpecEntry {
        pattern: "word/numbering.xml",
        content_type: Some(content_type::WORD_NUMBERING),
        expected_parser: "DocxParser::parse_numbering",
    },
    PartSpecEntry {
        pattern: "word/comments.xml",
        content_type: Some(content_type::WORD_COMMENTS),
        expected_parser: "DocxParser::parse_comments",
    },
    PartSpecEntry {
        pattern: "word/commentsExtended.xml",
        content_type: Some(content_type::WORD_COMMENTS_EXTENDED),
        expected_parser: "DocxParser::parse_comments_extended",
    },
    PartSpecEntry {
        pattern: "word/commentsIds.xml",
        content_type: Some(content_type::WORD_COMMENTS_IDS),
        expected_parser: "DocxParser::parse_comments_ids",
    },
    PartSpecEntry {
        pattern: "word/footnotes.xml",
        content_type: Some(content_type::WORD_FOOTNOTES),
        expected_parser: "DocxParser::parse_notes",
    },
    PartSpecEntry {
        pattern: "word/endnotes.xml",
        content_type: Some(content_type::WORD_ENDNOTES),
        expected_parser: "DocxParser::parse_notes",
    },
    PartSpecEntry {
        pattern: "word/header*.xml",
        content_type: Some(content_type::WORD_HEADER),
        expected_parser: "DocxParser::parse_header_footer",
    },
    PartSpecEntry {
        pattern: "word/footer*.xml",
        content_type: Some(content_type::WORD_FOOTER),
        expected_parser: "DocxParser::parse_header_footer",
    },
    PartSpecEntry {
        pattern: "word/settings.xml",
        content_type: Some(content_type::WORD_SETTINGS),
        expected_parser: "DocxParser::parse_settings",
    },
    PartSpecEntry {
        pattern: "word/webSettings.xml",
        content_type: Some(content_type::WORD_WEB_SETTINGS),
        expected_parser: "DocxParser::parse_web_settings",
    },
    PartSpecEntry {
        pattern: "word/fontTable.xml",
        content_type: Some(content_type::WORD_FONT_TABLE),
        expected_parser: "DocxParser::parse_font_table",
    },
    PartSpecEntry {
        pattern: "word/people.xml",
        content_type: Some(content_type::WORD_PEOPLE),
        expected_parser: "parse_people_part",
    },
    PartSpecEntry {
        pattern: "word/theme/theme*.xml",
        content_type: Some(content_type::WORD_THEME),
        expected_parser: "parse_theme",
    },
    PartSpecEntry {
        pattern: "word/drawings/*.xml",
        content_type: Some(content_type::WORD_DRAWING),
        expected_parser: "parse_drawingml",
    },
    PartSpecEntry {
        pattern: "word/charts/*.xml",
        content_type: Some(content_type::WORD_CHART),
        expected_parser: "parse_chart_data",
    },
    PartSpecEntry {
        pattern: "word/diagrams/data*.xml",
        content_type: Some(content_type::WORD_DIAGRAM_DATA),
        expected_parser: "parse_smartart_data",
    },
    PartSpecEntry {
        pattern: "word/diagrams/layout*.xml",
        content_type: Some(content_type::WORD_DIAGRAM_LAYOUT),
        expected_parser: "parse_smartart_layout",
    },
    PartSpecEntry {
        pattern: "word/diagrams/colors*.xml",
        content_type: Some(content_type::WORD_DIAGRAM_COLORS),
        expected_parser: "parse_smartart_colors",
    },
    PartSpecEntry {
        pattern: "word/diagrams/quickStyle*.xml",
        content_type: Some(content_type::WORD_DIAGRAM_STYLE),
        expected_parser: "parse_smartart_style",
    },
    PartSpecEntry {
        pattern: "word/media/*",
        content_type: None,
        expected_parser: "parse_shared_parts(media)",
    },
    PartSpecEntry {
        pattern: "word/embeddings/*.bin",
        content_type: Some(content_type::OLE_OBJECT),
        expected_parser: "parse_shared_parts(ole)",
    },
    PartSpecEntry {
        pattern: "word/vmlDrawing*.vml",
        content_type: Some(content_type::WORD_VML_DRAWING),
        expected_parser: "parse_vml_drawing",
    },
    PartSpecEntry {
        pattern: "word/comments*.xml.rels",
        content_type: Some(content_type::RELATIONSHIPS),
        expected_parser: "Relationships::parse",
    },
    PartSpecEntry {
        pattern: "word/footnotes.xml.rels",
        content_type: Some(content_type::RELATIONSHIPS),
        expected_parser: "Relationships::parse",
    },
    PartSpecEntry {
        pattern: "word/endnotes.xml.rels",
        content_type: Some(content_type::RELATIONSHIPS),
        expected_parser: "Relationships::parse",
    },
    PartSpecEntry {
        pattern: "word/vbaProject.bin",
        content_type: Some(content_type::VBA_PROJECT),
        expected_parser: "scan_security_content(vba)",
    },
    PartSpecEntry {
        pattern: "word/vbaData.xml",
        content_type: Some(content_type::WORD_VBA_DATA),
        expected_parser: "scan_security_content(vba)",
    },
    PartSpecEntry {
        pattern: "word/activeX/*.xml",
        content_type: Some(content_type::WORD_ACTIVEX_XML),
        expected_parser: "scan_security_content(activex)",
    },
    PartSpecEntry {
        pattern: "word/activeX/*.bin",
        content_type: Some(content_type::WORD_ACTIVEX_BIN),
        expected_parser: "scan_security_content(activex)",
    },
    PartSpecEntry {
        pattern: "word/webExtensions/webExtension*.xml",
        content_type: Some(content_type::WORD_WEB_EXTENSION),
        expected_parser: "parse_web_extensions",
    },
    PartSpecEntry {
        pattern: "word/webExtensions/taskpanes.xml",
        content_type: Some(content_type::WORD_WEB_EXTENSION_TASKPANES),
        expected_parser: "parse_web_extension_taskpanes",
    },
];

const SPREADSHEET_PARTS: &[PartSpecEntry] = &[
    PartSpecEntry {
        pattern: "_rels/.rels",
        content_type: Some(content_type::RELATIONSHIPS),
        expected_parser: "Relationships::parse",
    },
    PartSpecEntry {
        pattern: "xl/_rels/*.rels",
        content_type: Some(content_type::RELATIONSHIPS),
        expected_parser: "Relationships::parse",
    },
    PartSpecEntry {
        pattern: "docProps/core.xml",
        content_type: Some(content_type::CORE_PROPERTIES),
        expected_parser: "parse_metadata(core)",
    },
    PartSpecEntry {
        pattern: "docProps/app.xml",
        content_type: Some(content_type::EXTENDED_PROPERTIES),
        expected_parser: "parse_metadata(app)",
    },
    PartSpecEntry {
        pattern: "docProps/custom.xml",
        content_type: Some(content_type::CUSTOM_PROPERTIES),
        expected_parser: "parse_custom_properties",
    },
    PartSpecEntry {
        pattern: "customXml/*.xml",
        content_type: Some(content_type::CUSTOM_XML),
        expected_parser: "parse_custom_xml_part",
    },
    PartSpecEntry {
        pattern: "customXml/itemProps*.xml",
        content_type: Some(content_type::CUSTOM_XML_PROPERTIES),
        expected_parser: "parse_custom_xml_part",
    },
    PartSpecEntry {
        pattern: "_xmlsignatures/*.xml",
        content_type: Some(content_type::DIGITAL_SIGNATURE_XML),
        expected_parser: "parse_shared_parts(signature)",
    },
    PartSpecEntry {
        pattern: "_xmlsignatures/origin.sigs",
        content_type: Some(content_type::DIGITAL_SIGNATURE_ORIGIN),
        expected_parser: "parse_shared_parts(signature)",
    },
    PartSpecEntry {
        pattern: "xl/workbook.xml",
        content_type: Some(content_type::EXCEL_WORKBOOK),
        expected_parser: "XlsxParser::parse_workbook",
    },
    PartSpecEntry {
        pattern: "xl/workbook.xml",
        content_type: Some(content_type::EXCEL_WORKBOOK_MACRO),
        expected_parser: "XlsxParser::parse_workbook",
    },
    PartSpecEntry {
        pattern: "xl/workbook.bin",
        content_type: Some(content_type::EXCEL_WORKBOOK_BIN),
        expected_parser: "OoxmlParser::parse_xlsb",
    },
    PartSpecEntry {
        pattern: "xl/worksheets/*.xml",
        content_type: Some(content_type::EXCEL_WORKSHEET),
        expected_parser: "XlsxParser::parse_worksheet",
    },
    PartSpecEntry {
        pattern: "xl/chartsheets/*.xml",
        content_type: Some(content_type::EXCEL_CHARTSHEET),
        expected_parser: "XlsxParser::parse_chartsheet",
    },
    PartSpecEntry {
        pattern: "xl/dialogsheets/*.xml",
        content_type: Some(content_type::EXCEL_DIALOGSHEET),
        expected_parser: "XlsxParser::parse_worksheet",
    },
    PartSpecEntry {
        pattern: "xl/macrosheets/*.xml",
        content_type: Some(content_type::EXCEL_MACROSHEET),
        expected_parser: "XlsxParser::parse_worksheet",
    },
    PartSpecEntry {
        pattern: "xl/sharedStrings.xml",
        content_type: Some(content_type::EXCEL_SHARED_STRINGS),
        expected_parser: "parse_shared_strings_table",
    },
    PartSpecEntry {
        pattern: "xl/styles.xml",
        content_type: Some(content_type::EXCEL_STYLES),
        expected_parser: "parse_styles",
    },
    PartSpecEntry {
        pattern: "xl/calcChain.xml",
        content_type: Some(content_type::EXCEL_CALC_CHAIN),
        expected_parser: "parse_calc_chain",
    },
    PartSpecEntry {
        pattern: "xl/comments*.xml",
        content_type: Some(content_type::EXCEL_COMMENTS),
        expected_parser: "parse_sheet_comments",
    },
    PartSpecEntry {
        pattern: "xl/threadedComments/*.xml",
        content_type: Some(content_type::EXCEL_THREADED_COMMENTS),
        expected_parser: "parse_threaded_comments",
    },
    PartSpecEntry {
        pattern: "xl/drawings/*.xml",
        content_type: Some(content_type::EXCEL_DRAWING),
        expected_parser: "XlsxParser::parse_drawing",
    },
    PartSpecEntry {
        pattern: "xl/charts/*.xml",
        content_type: Some(content_type::EXCEL_CHART),
        expected_parser: "XlsxParser::parse_chart",
    },
    PartSpecEntry {
        pattern: "xl/pivotTables/*.xml",
        content_type: Some(content_type::EXCEL_PIVOT_TABLE),
        expected_parser: "parse_pivot_table_definition",
    },
    PartSpecEntry {
        pattern: "xl/pivotCache/pivotCacheDefinition*.xml",
        content_type: Some(content_type::EXCEL_PIVOT_CACHE_DEF),
        expected_parser: "parse_pivot_cache",
    },
    PartSpecEntry {
        pattern: "xl/pivotCache/pivotCacheRecords*.xml",
        content_type: Some(content_type::EXCEL_PIVOT_CACHE_RECORDS),
        expected_parser: "parse_pivot_cache_records",
    },
    PartSpecEntry {
        pattern: "xl/tables/*.xml",
        content_type: Some(content_type::EXCEL_TABLE),
        expected_parser: "parse_table_definition",
    },
    PartSpecEntry {
        pattern: "xl/queryTables/*.xml",
        content_type: Some(content_type::EXCEL_QUERY_TABLE),
        expected_parser: "parse_query_table",
    },
    PartSpecEntry {
        pattern: "xl/connections.xml",
        content_type: Some(content_type::EXCEL_CONNECTIONS),
        expected_parser: "parse_connections_part",
    },
    PartSpecEntry {
        pattern: "xl/externalLinks/*.xml",
        content_type: Some(content_type::EXCEL_EXTERNAL_LINK),
        expected_parser: "parse_external_links",
    },
    PartSpecEntry {
        pattern: "xl/metadata.xml",
        content_type: Some(content_type::EXCEL_SHEET_METADATA),
        expected_parser: "parse_sheet_metadata",
    },
    PartSpecEntry {
        pattern: "xl/embeddings/*.bin",
        content_type: Some(content_type::OLE_OBJECT),
        expected_parser: "parse_shared_parts(ole)",
    },
    PartSpecEntry {
        pattern: "xl/persons/person.xml",
        content_type: Some(content_type::EXCEL_PERSON),
        expected_parser: "parse_person_part",
    },
    PartSpecEntry {
        pattern: "xl/slicers/*.xml",
        content_type: Some(content_type::EXCEL_SLICER),
        expected_parser: "parse_slicers",
    },
    PartSpecEntry {
        pattern: "xl/timelines/*.xml",
        content_type: Some(content_type::EXCEL_TIMELINE),
        expected_parser: "parse_timelines",
    },
    PartSpecEntry {
        pattern: "xl/activeX/*.xml",
        content_type: Some(content_type::EXCEL_ACTIVEX_XML),
        expected_parser: "scan_security_content(activex)",
    },
    PartSpecEntry {
        pattern: "xl/activeX/*.bin",
        content_type: Some(content_type::EXCEL_ACTIVEX_BIN),
        expected_parser: "scan_security_content(activex)",
    },
    PartSpecEntry {
        pattern: "xl/vbaProject.bin",
        content_type: Some(content_type::VBA_PROJECT),
        expected_parser: "scan_security_content(vba)",
    },
];

const PRESENTATION_PARTS: &[PartSpecEntry] = &[
    PartSpecEntry {
        pattern: "_rels/.rels",
        content_type: Some(content_type::RELATIONSHIPS),
        expected_parser: "Relationships::parse",
    },
    PartSpecEntry {
        pattern: "ppt/_rels/*.rels",
        content_type: Some(content_type::RELATIONSHIPS),
        expected_parser: "Relationships::parse",
    },
    PartSpecEntry {
        pattern: "docProps/core.xml",
        content_type: Some(content_type::CORE_PROPERTIES),
        expected_parser: "parse_metadata(core)",
    },
    PartSpecEntry {
        pattern: "docProps/app.xml",
        content_type: Some(content_type::EXTENDED_PROPERTIES),
        expected_parser: "parse_metadata(app)",
    },
    PartSpecEntry {
        pattern: "docProps/custom.xml",
        content_type: Some(content_type::CUSTOM_PROPERTIES),
        expected_parser: "parse_custom_properties",
    },
    PartSpecEntry {
        pattern: "customXml/*.xml",
        content_type: Some(content_type::CUSTOM_XML),
        expected_parser: "parse_custom_xml_part",
    },
    PartSpecEntry {
        pattern: "customXml/itemProps*.xml",
        content_type: Some(content_type::CUSTOM_XML_PROPERTIES),
        expected_parser: "parse_custom_xml_part",
    },
    PartSpecEntry {
        pattern: "_xmlsignatures/*.xml",
        content_type: Some(content_type::DIGITAL_SIGNATURE_XML),
        expected_parser: "parse_shared_parts(signature)",
    },
    PartSpecEntry {
        pattern: "_xmlsignatures/origin.sigs",
        content_type: Some(content_type::DIGITAL_SIGNATURE_ORIGIN),
        expected_parser: "parse_shared_parts(signature)",
    },
    PartSpecEntry {
        pattern: "ppt/presentation.xml",
        content_type: Some(content_type::PPTX_PRESENTATION),
        expected_parser: "PptxParser::parse_presentation",
    },
    PartSpecEntry {
        pattern: "ppt/presentation.xml",
        content_type: Some(content_type::PPTX_PRESENTATION_MACRO),
        expected_parser: "PptxParser::parse_presentation",
    },
    PartSpecEntry {
        pattern: "ppt/presentation.xml",
        content_type: Some(content_type::PPTX_PRESENTATION),
        expected_parser: "parse_presentation_info",
    },
    PartSpecEntry {
        pattern: "ppt/slides/*.xml",
        content_type: Some(content_type::PPTX_SLIDE),
        expected_parser: "PptxParser::parse_slide",
    },
    PartSpecEntry {
        pattern: "ppt/presProps.xml",
        content_type: Some(content_type::PPTX_PRES_PROPS),
        expected_parser: "parse_presentation_properties",
    },
    PartSpecEntry {
        pattern: "ppt/viewProps.xml",
        content_type: Some(content_type::PPTX_VIEW_PROPS),
        expected_parser: "parse_view_properties",
    },
    PartSpecEntry {
        pattern: "ppt/tableStyles.xml",
        content_type: Some(content_type::PPTX_TABLE_STYLES),
        expected_parser: "parse_table_styles",
    },
    PartSpecEntry {
        pattern: "ppt/slideMasters/*.xml",
        content_type: Some(content_type::PPTX_SLIDE_MASTER),
        expected_parser: "parse_slide_master",
    },
    PartSpecEntry {
        pattern: "ppt/slideLayouts/*.xml",
        content_type: Some(content_type::PPTX_SLIDE_LAYOUT),
        expected_parser: "parse_slide_layout",
    },
    PartSpecEntry {
        pattern: "ppt/notesMasters/*.xml",
        content_type: Some(content_type::PPTX_NOTES_MASTER),
        expected_parser: "parse_notes_master",
    },
    PartSpecEntry {
        pattern: "ppt/handoutMasters/*.xml",
        content_type: Some(content_type::PPTX_HANDOUT_MASTER),
        expected_parser: "parse_handout_master",
    },
    PartSpecEntry {
        pattern: "ppt/notesSlides/*.xml",
        content_type: Some(content_type::PPTX_NOTES_SLIDE),
        expected_parser: "parse_notes_slide",
    },
    PartSpecEntry {
        pattern: "ppt/diagrams/data*.xml",
        content_type: Some(content_type::PPTX_DIAGRAM_DATA),
        expected_parser: "parse_smartart_data",
    },
    PartSpecEntry {
        pattern: "ppt/diagrams/layout*.xml",
        content_type: Some(content_type::PPTX_DIAGRAM_LAYOUT),
        expected_parser: "parse_smartart_layout",
    },
    PartSpecEntry {
        pattern: "ppt/diagrams/colors*.xml",
        content_type: Some(content_type::PPTX_DIAGRAM_COLORS),
        expected_parser: "parse_smartart_colors",
    },
    PartSpecEntry {
        pattern: "ppt/diagrams/quickStyle*.xml",
        content_type: Some(content_type::PPTX_DIAGRAM_STYLE),
        expected_parser: "parse_smartart_style",
    },
    PartSpecEntry {
        pattern: "ppt/charts/*.xml",
        content_type: Some(content_type::PPTX_CHART),
        expected_parser: "parse_chart_data",
    },
    PartSpecEntry {
        pattern: "ppt/media/*",
        content_type: None,
        expected_parser: "parse_shared_parts(media)",
    },
    PartSpecEntry {
        pattern: "ppt/embeddings/*.bin",
        content_type: Some(content_type::OLE_OBJECT),
        expected_parser: "parse_shared_parts(ole)",
    },
    PartSpecEntry {
        pattern: "ppt/commentAuthors.xml",
        content_type: Some(content_type::PPTX_COMMENT_AUTHORS),
        expected_parser: "parse_comment_authors",
    },
    PartSpecEntry {
        pattern: "ppt/comments/*.xml",
        content_type: Some(content_type::PPTX_COMMENTS),
        expected_parser: "parse_comments",
    },
    PartSpecEntry {
        pattern: "ppt/tags/*.xml",
        content_type: Some(content_type::PPTX_TAGS),
        expected_parser: "parse_presentation_tags",
    },
    PartSpecEntry {
        pattern: "ppt/people.xml",
        content_type: Some(content_type::PPTX_PEOPLE),
        expected_parser: "parse_people_part",
    },
    PartSpecEntry {
        pattern: "ppt/activeX/*.xml",
        content_type: Some(content_type::PPTX_ACTIVEX_XML),
        expected_parser: "scan_security_content(activex)",
    },
    PartSpecEntry {
        pattern: "ppt/activeX/*.bin",
        content_type: Some(content_type::PPTX_ACTIVEX_BIN),
        expected_parser: "scan_security_content(activex)",
    },
    PartSpecEntry {
        pattern: "ppt/vbaProject.bin",
        content_type: Some(content_type::VBA_PROJECT),
        expected_parser: "scan_security_content(vba)",
    },
];

pub fn registry_for(format: DocumentFormat) -> Vec<PartSpec> {
    match format {
        DocumentFormat::WordProcessing => build_registry(format, WORD_PARTS),
        DocumentFormat::Spreadsheet => build_registry(format, SPREADSHEET_PARTS),
        DocumentFormat::Presentation => build_registry(format, PRESENTATION_PARTS),
        DocumentFormat::OdfText
        | DocumentFormat::OdfSpreadsheet
        | DocumentFormat::OdfPresentation
        | DocumentFormat::Hwp
        | DocumentFormat::Hwpx
        | DocumentFormat::Rtf => Vec::new(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_registry_includes_macro_projects() {
        let xlsx = registry_for(DocumentFormat::Spreadsheet);
        assert!(xlsx.iter().any(|p| p.pattern == "xl/vbaProject.bin"));

        let pptx = registry_for(DocumentFormat::Presentation);
        assert!(pptx.iter().any(|p| p.pattern == "ppt/vbaProject.bin"));
    }

    #[test]
    fn test_registry_includes_package_parts_and_rels() {
        let docx = registry_for(DocumentFormat::WordProcessing);
        assert!(docx.iter().any(|p| p.pattern == "_rels/.rels"));
        assert!(docx.iter().any(|p| p.pattern == "word/_rels/*.rels"));
        assert!(docx.iter().any(|p| p.pattern == "docProps/core.xml"));
        assert!(docx.iter().any(|p| p.pattern == "docProps/app.xml"));
        assert!(docx.iter().any(|p| p.pattern == "docProps/custom.xml"));
        assert!(docx.iter().any(|p| p.pattern == "customXml/*.xml"));
        assert!(docx.iter().any(|p| p.pattern == "customXml/itemProps*.xml"));
    }

    #[test]
    fn test_registry_includes_signature_and_customxml_props() {
        let docx = registry_for(DocumentFormat::WordProcessing);
        assert!(docx.iter().any(|p| p.pattern == "_xmlsignatures/*.xml"));
        assert!(docx
            .iter()
            .any(|p| p.pattern == "_xmlsignatures/origin.sigs"));
        assert!(docx.iter().any(|p| p.pattern == "customXml/itemProps*.xml"));
    }
}
