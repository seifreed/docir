//! OOXML Part Registry for coverage tracking.

use crate::ooxml::content_types::content_type;
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

pub fn registry_for(format: DocumentFormat) -> Vec<PartSpec> {
    let mut entries: Vec<PartSpec> = Vec::new();

    match format {
        DocumentFormat::WordProcessing => {
            entries.extend([
                PartSpec {
                    format,
                    pattern: "_rels/.rels",
                    content_type: Some(content_type::RELATIONSHIPS),
                    expected_parser: "Relationships::parse",
                },
                PartSpec {
                    format,
                    pattern: "word/_rels/*.rels",
                    content_type: Some(content_type::RELATIONSHIPS),
                    expected_parser: "Relationships::parse",
                },
                PartSpec {
                    format,
                    pattern: "docProps/core.xml",
                    content_type: Some(content_type::CORE_PROPERTIES),
                    expected_parser: "parse_metadata(core)",
                },
                PartSpec {
                    format,
                    pattern: "docProps/app.xml",
                    content_type: Some(content_type::EXTENDED_PROPERTIES),
                    expected_parser: "parse_metadata(app)",
                },
                PartSpec {
                    format,
                    pattern: "docProps/custom.xml",
                    content_type: Some(content_type::CUSTOM_PROPERTIES),
                    expected_parser: "parse_custom_properties",
                },
                PartSpec {
                    format,
                    pattern: "customXml/*.xml",
                    content_type: Some(content_type::CUSTOM_XML),
                    expected_parser: "parse_custom_xml_part",
                },
                PartSpec {
                    format,
                    pattern: "customXml/itemProps*.xml",
                    content_type: Some(content_type::CUSTOM_XML_PROPERTIES),
                    expected_parser: "parse_custom_xml_part",
                },
                PartSpec {
                    format,
                    pattern: "_xmlsignatures/*.xml",
                    content_type: Some(content_type::DIGITAL_SIGNATURE_XML),
                    expected_parser: "parse_shared_parts(signature)",
                },
                PartSpec {
                    format,
                    pattern: "_xmlsignatures/origin.sigs",
                    content_type: Some(content_type::DIGITAL_SIGNATURE_ORIGIN),
                    expected_parser: "parse_shared_parts(signature)",
                },
                PartSpec {
                    format,
                    pattern: "word/document.xml",
                    content_type: Some(content_type::WORD_DOCUMENT),
                    expected_parser: "DocxParser::parse_document",
                },
                PartSpec {
                    format,
                    pattern: "word/document.xml",
                    content_type: Some(content_type::WORD_DOCUMENT_MACRO),
                    expected_parser: "DocxParser::parse_document",
                },
                PartSpec {
                    format,
                    pattern: "word/glossary/document.xml",
                    content_type: Some(content_type::WORD_GLOSSARY),
                    expected_parser: "DocxParser::parse_glossary_document",
                },
                PartSpec {
                    format,
                    pattern: "word/styles.xml",
                    content_type: Some(content_type::WORD_STYLES),
                    expected_parser: "DocxParser::parse_styles",
                },
                PartSpec {
                    format,
                    pattern: "word/numbering.xml",
                    content_type: Some(content_type::WORD_NUMBERING),
                    expected_parser: "DocxParser::parse_numbering",
                },
                PartSpec {
                    format,
                    pattern: "word/comments.xml",
                    content_type: Some(content_type::WORD_COMMENTS),
                    expected_parser: "DocxParser::parse_comments",
                },
                PartSpec {
                    format,
                    pattern: "word/commentsExtended.xml",
                    content_type: Some(content_type::WORD_COMMENTS_EXTENDED),
                    expected_parser: "DocxParser::parse_comments_extended",
                },
                PartSpec {
                    format,
                    pattern: "word/commentsIds.xml",
                    content_type: Some(content_type::WORD_COMMENTS_IDS),
                    expected_parser: "DocxParser::parse_comments_ids",
                },
                PartSpec {
                    format,
                    pattern: "word/footnotes.xml",
                    content_type: Some(content_type::WORD_FOOTNOTES),
                    expected_parser: "DocxParser::parse_notes",
                },
                PartSpec {
                    format,
                    pattern: "word/endnotes.xml",
                    content_type: Some(content_type::WORD_ENDNOTES),
                    expected_parser: "DocxParser::parse_notes",
                },
                PartSpec {
                    format,
                    pattern: "word/header*.xml",
                    content_type: Some(content_type::WORD_HEADER),
                    expected_parser: "DocxParser::parse_header_footer",
                },
                PartSpec {
                    format,
                    pattern: "word/footer*.xml",
                    content_type: Some(content_type::WORD_FOOTER),
                    expected_parser: "DocxParser::parse_header_footer",
                },
                PartSpec {
                    format,
                    pattern: "word/settings.xml",
                    content_type: Some(content_type::WORD_SETTINGS),
                    expected_parser: "DocxParser::parse_settings",
                },
                PartSpec {
                    format,
                    pattern: "word/webSettings.xml",
                    content_type: Some(content_type::WORD_WEB_SETTINGS),
                    expected_parser: "DocxParser::parse_web_settings",
                },
                PartSpec {
                    format,
                    pattern: "word/fontTable.xml",
                    content_type: Some(content_type::WORD_FONT_TABLE),
                    expected_parser: "DocxParser::parse_font_table",
                },
                PartSpec {
                    format,
                    pattern: "word/people.xml",
                    content_type: Some(content_type::WORD_PEOPLE),
                    expected_parser: "parse_people_part",
                },
                PartSpec {
                    format,
                    pattern: "word/theme/theme*.xml",
                    content_type: Some(content_type::WORD_THEME),
                    expected_parser: "parse_theme",
                },
                PartSpec {
                    format,
                    pattern: "word/drawings/*.xml",
                    content_type: Some(content_type::WORD_DRAWING),
                    expected_parser: "parse_drawingml",
                },
                PartSpec {
                    format,
                    pattern: "word/charts/*.xml",
                    content_type: Some(content_type::WORD_CHART),
                    expected_parser: "parse_chart_data",
                },
                PartSpec {
                    format,
                    pattern: "word/diagrams/data*.xml",
                    content_type: Some(content_type::WORD_DIAGRAM_DATA),
                    expected_parser: "parse_smartart_data",
                },
                PartSpec {
                    format,
                    pattern: "word/diagrams/layout*.xml",
                    content_type: Some(content_type::WORD_DIAGRAM_LAYOUT),
                    expected_parser: "parse_smartart_layout",
                },
                PartSpec {
                    format,
                    pattern: "word/diagrams/colors*.xml",
                    content_type: Some(content_type::WORD_DIAGRAM_COLORS),
                    expected_parser: "parse_smartart_colors",
                },
                PartSpec {
                    format,
                    pattern: "word/diagrams/quickStyle*.xml",
                    content_type: Some(content_type::WORD_DIAGRAM_STYLE),
                    expected_parser: "parse_smartart_style",
                },
                PartSpec {
                    format,
                    pattern: "word/media/*",
                    content_type: None,
                    expected_parser: "parse_shared_parts(media)",
                },
                PartSpec {
                    format,
                    pattern: "word/embeddings/*.bin",
                    content_type: Some(content_type::OLE_OBJECT),
                    expected_parser: "parse_shared_parts(ole)",
                },
                PartSpec {
                    format,
                    pattern: "word/vmlDrawing*.vml",
                    content_type: Some(content_type::WORD_VML_DRAWING),
                    expected_parser: "parse_vml_drawing",
                },
                PartSpec {
                    format,
                    pattern: "word/comments*.xml.rels",
                    content_type: Some(content_type::RELATIONSHIPS),
                    expected_parser: "Relationships::parse",
                },
                PartSpec {
                    format,
                    pattern: "word/footnotes.xml.rels",
                    content_type: Some(content_type::RELATIONSHIPS),
                    expected_parser: "Relationships::parse",
                },
                PartSpec {
                    format,
                    pattern: "word/endnotes.xml.rels",
                    content_type: Some(content_type::RELATIONSHIPS),
                    expected_parser: "Relationships::parse",
                },
                PartSpec {
                    format,
                    pattern: "word/vbaProject.bin",
                    content_type: Some(content_type::VBA_PROJECT),
                    expected_parser: "scan_security_content(vba)",
                },
                PartSpec {
                    format,
                    pattern: "word/vbaData.xml",
                    content_type: Some(content_type::WORD_VBA_DATA),
                    expected_parser: "scan_security_content(vba)",
                },
                PartSpec {
                    format,
                    pattern: "word/activeX/*.xml",
                    content_type: Some(content_type::WORD_ACTIVEX_XML),
                    expected_parser: "scan_security_content(activex)",
                },
                PartSpec {
                    format,
                    pattern: "word/activeX/*.bin",
                    content_type: Some(content_type::WORD_ACTIVEX_BIN),
                    expected_parser: "scan_security_content(activex)",
                },
                PartSpec {
                    format,
                    pattern: "word/webExtensions/webExtension*.xml",
                    content_type: Some(content_type::WORD_WEB_EXTENSION),
                    expected_parser: "parse_web_extensions",
                },
                PartSpec {
                    format,
                    pattern: "word/webExtensions/taskpanes.xml",
                    content_type: Some(content_type::WORD_WEB_EXTENSION_TASKPANES),
                    expected_parser: "parse_web_extension_taskpanes",
                },
            ]);
        }
        DocumentFormat::Spreadsheet => {
            entries.extend([
                PartSpec {
                    format,
                    pattern: "_rels/.rels",
                    content_type: Some(content_type::RELATIONSHIPS),
                    expected_parser: "Relationships::parse",
                },
                PartSpec {
                    format,
                    pattern: "xl/_rels/*.rels",
                    content_type: Some(content_type::RELATIONSHIPS),
                    expected_parser: "Relationships::parse",
                },
                PartSpec {
                    format,
                    pattern: "docProps/core.xml",
                    content_type: Some(content_type::CORE_PROPERTIES),
                    expected_parser: "parse_metadata(core)",
                },
                PartSpec {
                    format,
                    pattern: "docProps/app.xml",
                    content_type: Some(content_type::EXTENDED_PROPERTIES),
                    expected_parser: "parse_metadata(app)",
                },
                PartSpec {
                    format,
                    pattern: "docProps/custom.xml",
                    content_type: Some(content_type::CUSTOM_PROPERTIES),
                    expected_parser: "parse_custom_properties",
                },
                PartSpec {
                    format,
                    pattern: "customXml/*.xml",
                    content_type: Some(content_type::CUSTOM_XML),
                    expected_parser: "parse_custom_xml_part",
                },
                PartSpec {
                    format,
                    pattern: "customXml/itemProps*.xml",
                    content_type: Some(content_type::CUSTOM_XML_PROPERTIES),
                    expected_parser: "parse_custom_xml_part",
                },
                PartSpec {
                    format,
                    pattern: "_xmlsignatures/*.xml",
                    content_type: Some(content_type::DIGITAL_SIGNATURE_XML),
                    expected_parser: "parse_shared_parts(signature)",
                },
                PartSpec {
                    format,
                    pattern: "_xmlsignatures/origin.sigs",
                    content_type: Some(content_type::DIGITAL_SIGNATURE_ORIGIN),
                    expected_parser: "parse_shared_parts(signature)",
                },
                PartSpec {
                    format,
                    pattern: "xl/workbook.xml",
                    content_type: Some(content_type::EXCEL_WORKBOOK),
                    expected_parser: "XlsxParser::parse_workbook",
                },
                PartSpec {
                    format,
                    pattern: "xl/workbook.xml",
                    content_type: Some(content_type::EXCEL_WORKBOOK_MACRO),
                    expected_parser: "XlsxParser::parse_workbook",
                },
                PartSpec {
                    format,
                    pattern: "xl/workbook.bin",
                    content_type: Some(content_type::EXCEL_WORKBOOK_BIN),
                    expected_parser: "OoxmlParser::parse_xlsb",
                },
                PartSpec {
                    format,
                    pattern: "xl/worksheets/*.xml",
                    content_type: Some(content_type::EXCEL_WORKSHEET),
                    expected_parser: "XlsxParser::parse_worksheet",
                },
                PartSpec {
                    format,
                    pattern: "xl/chartsheets/*.xml",
                    content_type: Some(content_type::EXCEL_CHARTSHEET),
                    expected_parser: "XlsxParser::parse_chartsheet",
                },
                PartSpec {
                    format,
                    pattern: "xl/dialogsheets/*.xml",
                    content_type: Some(content_type::EXCEL_DIALOGSHEET),
                    expected_parser: "XlsxParser::parse_worksheet",
                },
                PartSpec {
                    format,
                    pattern: "xl/macrosheets/*.xml",
                    content_type: Some(content_type::EXCEL_MACROSHEET),
                    expected_parser: "XlsxParser::parse_worksheet",
                },
                PartSpec {
                    format,
                    pattern: "xl/sharedStrings.xml",
                    content_type: Some(content_type::EXCEL_SHARED_STRINGS),
                    expected_parser: "parse_shared_strings_table",
                },
                PartSpec {
                    format,
                    pattern: "xl/styles.xml",
                    content_type: Some(content_type::EXCEL_STYLES),
                    expected_parser: "parse_styles",
                },
                PartSpec {
                    format,
                    pattern: "xl/calcChain.xml",
                    content_type: Some(content_type::EXCEL_CALC_CHAIN),
                    expected_parser: "parse_calc_chain",
                },
                PartSpec {
                    format,
                    pattern: "xl/comments*.xml",
                    content_type: Some(content_type::EXCEL_COMMENTS),
                    expected_parser: "parse_sheet_comments",
                },
                PartSpec {
                    format,
                    pattern: "xl/threadedComments/*.xml",
                    content_type: Some(content_type::EXCEL_THREADED_COMMENTS),
                    expected_parser: "parse_threaded_comments",
                },
                PartSpec {
                    format,
                    pattern: "xl/drawings/*.xml",
                    content_type: Some(content_type::EXCEL_DRAWING),
                    expected_parser: "XlsxParser::parse_drawing",
                },
                PartSpec {
                    format,
                    pattern: "xl/charts/*.xml",
                    content_type: Some(content_type::EXCEL_CHART),
                    expected_parser: "XlsxParser::parse_chart",
                },
                PartSpec {
                    format,
                    pattern: "xl/pivotTables/*.xml",
                    content_type: Some(content_type::EXCEL_PIVOT_TABLE),
                    expected_parser: "parse_pivot_table_definition",
                },
                PartSpec {
                    format,
                    pattern: "xl/pivotCache/pivotCacheDefinition*.xml",
                    content_type: Some(content_type::EXCEL_PIVOT_CACHE_DEF),
                    expected_parser: "parse_pivot_cache",
                },
                PartSpec {
                    format,
                    pattern: "xl/pivotCache/pivotCacheRecords*.xml",
                    content_type: Some(content_type::EXCEL_PIVOT_CACHE_RECORDS),
                    expected_parser: "parse_pivot_cache_records",
                },
                PartSpec {
                    format,
                    pattern: "xl/tables/*.xml",
                    content_type: Some(content_type::EXCEL_TABLE),
                    expected_parser: "parse_table_definition",
                },
                PartSpec {
                    format,
                    pattern: "xl/queryTables/*.xml",
                    content_type: Some(content_type::EXCEL_QUERY_TABLE),
                    expected_parser: "parse_query_table",
                },
                PartSpec {
                    format,
                    pattern: "xl/connections.xml",
                    content_type: Some(content_type::EXCEL_CONNECTIONS),
                    expected_parser: "parse_connections_part",
                },
                PartSpec {
                    format,
                    pattern: "xl/externalLinks/*.xml",
                    content_type: Some(content_type::EXCEL_EXTERNAL_LINK),
                    expected_parser: "parse_external_links",
                },
                PartSpec {
                    format,
                    pattern: "xl/metadata.xml",
                    content_type: Some(content_type::EXCEL_SHEET_METADATA),
                    expected_parser: "parse_sheet_metadata",
                },
                PartSpec {
                    format,
                    pattern: "xl/embeddings/*.bin",
                    content_type: Some(content_type::OLE_OBJECT),
                    expected_parser: "parse_shared_parts(ole)",
                },
                PartSpec {
                    format,
                    pattern: "xl/persons/person.xml",
                    content_type: Some(content_type::EXCEL_PERSON),
                    expected_parser: "parse_person_part",
                },
                PartSpec {
                    format,
                    pattern: "xl/slicers/*.xml",
                    content_type: Some(content_type::EXCEL_SLICER),
                    expected_parser: "parse_slicers",
                },
                PartSpec {
                    format,
                    pattern: "xl/timelines/*.xml",
                    content_type: Some(content_type::EXCEL_TIMELINE),
                    expected_parser: "parse_timelines",
                },
                PartSpec {
                    format,
                    pattern: "xl/activeX/*.xml",
                    content_type: Some(content_type::EXCEL_ACTIVEX_XML),
                    expected_parser: "scan_security_content(activex)",
                },
                PartSpec {
                    format,
                    pattern: "xl/activeX/*.bin",
                    content_type: Some(content_type::EXCEL_ACTIVEX_BIN),
                    expected_parser: "scan_security_content(activex)",
                },
                PartSpec {
                    format,
                    pattern: "xl/vbaProject.bin",
                    content_type: Some(content_type::VBA_PROJECT),
                    expected_parser: "scan_security_content(vba)",
                },
            ]);
        }
        DocumentFormat::Presentation => {
            entries.extend([
                PartSpec {
                    format,
                    pattern: "_rels/.rels",
                    content_type: Some(content_type::RELATIONSHIPS),
                    expected_parser: "Relationships::parse",
                },
                PartSpec {
                    format,
                    pattern: "ppt/_rels/*.rels",
                    content_type: Some(content_type::RELATIONSHIPS),
                    expected_parser: "Relationships::parse",
                },
                PartSpec {
                    format,
                    pattern: "docProps/core.xml",
                    content_type: Some(content_type::CORE_PROPERTIES),
                    expected_parser: "parse_metadata(core)",
                },
                PartSpec {
                    format,
                    pattern: "docProps/app.xml",
                    content_type: Some(content_type::EXTENDED_PROPERTIES),
                    expected_parser: "parse_metadata(app)",
                },
                PartSpec {
                    format,
                    pattern: "docProps/custom.xml",
                    content_type: Some(content_type::CUSTOM_PROPERTIES),
                    expected_parser: "parse_custom_properties",
                },
                PartSpec {
                    format,
                    pattern: "customXml/*.xml",
                    content_type: Some(content_type::CUSTOM_XML),
                    expected_parser: "parse_custom_xml_part",
                },
                PartSpec {
                    format,
                    pattern: "customXml/itemProps*.xml",
                    content_type: Some(content_type::CUSTOM_XML_PROPERTIES),
                    expected_parser: "parse_custom_xml_part",
                },
                PartSpec {
                    format,
                    pattern: "_xmlsignatures/*.xml",
                    content_type: Some(content_type::DIGITAL_SIGNATURE_XML),
                    expected_parser: "parse_shared_parts(signature)",
                },
                PartSpec {
                    format,
                    pattern: "_xmlsignatures/origin.sigs",
                    content_type: Some(content_type::DIGITAL_SIGNATURE_ORIGIN),
                    expected_parser: "parse_shared_parts(signature)",
                },
                PartSpec {
                    format,
                    pattern: "ppt/presentation.xml",
                    content_type: Some(content_type::PPTX_PRESENTATION),
                    expected_parser: "PptxParser::parse_presentation",
                },
                PartSpec {
                    format,
                    pattern: "ppt/presentation.xml",
                    content_type: Some(content_type::PPTX_PRESENTATION_MACRO),
                    expected_parser: "PptxParser::parse_presentation",
                },
                PartSpec {
                    format,
                    pattern: "ppt/presentation.xml",
                    content_type: Some(content_type::PPTX_PRESENTATION),
                    expected_parser: "parse_presentation_info",
                },
                PartSpec {
                    format,
                    pattern: "ppt/slides/*.xml",
                    content_type: Some(content_type::PPTX_SLIDE),
                    expected_parser: "PptxParser::parse_slide",
                },
                PartSpec {
                    format,
                    pattern: "ppt/presProps.xml",
                    content_type: Some(content_type::PPTX_PRES_PROPS),
                    expected_parser: "parse_presentation_properties",
                },
                PartSpec {
                    format,
                    pattern: "ppt/viewProps.xml",
                    content_type: Some(content_type::PPTX_VIEW_PROPS),
                    expected_parser: "parse_view_properties",
                },
                PartSpec {
                    format,
                    pattern: "ppt/tableStyles.xml",
                    content_type: Some(content_type::PPTX_TABLE_STYLES),
                    expected_parser: "parse_table_styles",
                },
                PartSpec {
                    format,
                    pattern: "ppt/slideMasters/*.xml",
                    content_type: Some(content_type::PPTX_SLIDE_MASTER),
                    expected_parser: "parse_slide_master",
                },
                PartSpec {
                    format,
                    pattern: "ppt/slideLayouts/*.xml",
                    content_type: Some(content_type::PPTX_SLIDE_LAYOUT),
                    expected_parser: "parse_slide_layout",
                },
                PartSpec {
                    format,
                    pattern: "ppt/notesMasters/*.xml",
                    content_type: Some(content_type::PPTX_NOTES_MASTER),
                    expected_parser: "parse_notes_master",
                },
                PartSpec {
                    format,
                    pattern: "ppt/handoutMasters/*.xml",
                    content_type: Some(content_type::PPTX_HANDOUT_MASTER),
                    expected_parser: "parse_handout_master",
                },
                PartSpec {
                    format,
                    pattern: "ppt/notesSlides/*.xml",
                    content_type: Some(content_type::PPTX_NOTES_SLIDE),
                    expected_parser: "parse_notes_slide",
                },
                PartSpec {
                    format,
                    pattern: "ppt/diagrams/data*.xml",
                    content_type: Some(content_type::PPTX_DIAGRAM_DATA),
                    expected_parser: "parse_smartart_data",
                },
                PartSpec {
                    format,
                    pattern: "ppt/diagrams/layout*.xml",
                    content_type: Some(content_type::PPTX_DIAGRAM_LAYOUT),
                    expected_parser: "parse_smartart_layout",
                },
                PartSpec {
                    format,
                    pattern: "ppt/diagrams/colors*.xml",
                    content_type: Some(content_type::PPTX_DIAGRAM_COLORS),
                    expected_parser: "parse_smartart_colors",
                },
                PartSpec {
                    format,
                    pattern: "ppt/diagrams/quickStyle*.xml",
                    content_type: Some(content_type::PPTX_DIAGRAM_STYLE),
                    expected_parser: "parse_smartart_style",
                },
                PartSpec {
                    format,
                    pattern: "ppt/charts/*.xml",
                    content_type: Some(content_type::PPTX_CHART),
                    expected_parser: "parse_chart_data",
                },
                PartSpec {
                    format,
                    pattern: "ppt/media/*",
                    content_type: None,
                    expected_parser: "parse_shared_parts(media)",
                },
                PartSpec {
                    format,
                    pattern: "ppt/embeddings/*.bin",
                    content_type: Some(content_type::OLE_OBJECT),
                    expected_parser: "parse_shared_parts(ole)",
                },
                PartSpec {
                    format,
                    pattern: "ppt/commentAuthors.xml",
                    content_type: Some(content_type::PPTX_COMMENT_AUTHORS),
                    expected_parser: "parse_comment_authors",
                },
                PartSpec {
                    format,
                    pattern: "ppt/comments/*.xml",
                    content_type: Some(content_type::PPTX_COMMENTS),
                    expected_parser: "parse_comments",
                },
                PartSpec {
                    format,
                    pattern: "ppt/tags/*.xml",
                    content_type: Some(content_type::PPTX_TAGS),
                    expected_parser: "parse_presentation_tags",
                },
                PartSpec {
                    format,
                    pattern: "ppt/people.xml",
                    content_type: Some(content_type::PPTX_PEOPLE),
                    expected_parser: "parse_people_part",
                },
                PartSpec {
                    format,
                    pattern: "ppt/activeX/*.xml",
                    content_type: Some(content_type::PPTX_ACTIVEX_XML),
                    expected_parser: "scan_security_content(activex)",
                },
                PartSpec {
                    format,
                    pattern: "ppt/activeX/*.bin",
                    content_type: Some(content_type::PPTX_ACTIVEX_BIN),
                    expected_parser: "scan_security_content(activex)",
                },
                PartSpec {
                    format,
                    pattern: "ppt/vbaProject.bin",
                    content_type: Some(content_type::VBA_PROJECT),
                    expected_parser: "scan_security_content(vba)",
                },
            ]);
        }
        DocumentFormat::OdfText
        | DocumentFormat::OdfSpreadsheet
        | DocumentFormat::OdfPresentation
        | DocumentFormat::Hwp
        | DocumentFormat::Hwpx
        | DocumentFormat::Rtf => {}
    }

    entries
}

fn matches_pattern(path: &str, pattern: &str) -> bool {
    if !pattern.contains('*') {
        return path == pattern;
    }

    let mut parts = pattern.splitn(2, '*');
    let prefix = parts.next().unwrap_or("");
    let suffix = parts.next().unwrap_or("");

    if !prefix.is_empty() && !path.starts_with(prefix) {
        return false;
    }
    if !suffix.is_empty() && !path.ends_with(suffix) {
        return false;
    }
    true
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
