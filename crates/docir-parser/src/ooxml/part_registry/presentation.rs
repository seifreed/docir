use super::PartSpecEntry;
use crate::ooxml::content_types::content_type;

pub(super) const PRESENTATION_PARTS: &[PartSpecEntry] = &[
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
