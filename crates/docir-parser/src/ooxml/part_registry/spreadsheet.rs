use super::PartSpecEntry;
use crate::ooxml::content_types::content_type;

pub(super) const SPREADSHEET_PARTS: &[PartSpecEntry] = &[
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
