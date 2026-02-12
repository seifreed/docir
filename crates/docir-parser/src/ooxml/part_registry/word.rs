use super::PartSpecEntry;
use crate::ooxml::content_types::content_type;

pub(super) const WORD_PARTS: &[PartSpecEntry] = &[
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
