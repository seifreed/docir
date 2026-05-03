use crate::io_support::with_file_bytes;
use crate::{AppResult, ParserConfig};
use docir_parser::ole::Cfb;
use docir_parser::ParseError as ParserParseError;
use serde::Serialize;
use std::path::Path;

const SUMMARY_INFO_STREAM: &str = "\u{0005}SummaryInformation";
const DOC_SUMMARY_INFO_STREAM: &str = "\u{0005}DocumentSummaryInformation";

/// Structured metadata extracted from classic OLE property-set streams.
#[derive(Debug, Clone, Serialize)]
pub struct MetadataInspection {
    pub container: String,
    pub section_count: usize,
    pub sections: Vec<MetadataSection>,
}

/// One logical property-set section.
#[derive(Debug, Clone, Serialize)]
pub struct MetadataSection {
    pub name: String,
    pub path: String,
    pub property_count: usize,
    pub properties: Vec<MetadataProperty>,
}

/// One extracted metadata property.
#[derive(Debug, Clone, Serialize)]
pub struct MetadataProperty {
    pub id: u32,
    pub name: String,
    pub value_type: String,
    pub value: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub display_value: Option<String>,
}

/// Inspect metadata from a legacy CFB/OLE file on disk.
pub fn inspect_metadata_path<P: AsRef<Path>>(
    path: P,
    config: &ParserConfig,
) -> AppResult<MetadataInspection> {
    with_file_bytes(path, config.max_input_size, inspect_metadata_bytes)
}

/// Inspect metadata from raw CFB/OLE bytes.
pub fn inspect_metadata_bytes(data: &[u8]) -> AppResult<MetadataInspection> {
    let cfb = Cfb::parse(data.to_vec())?;
    let mut sections = Vec::new();

    if let Some(bytes) = cfb.read_stream(SUMMARY_INFO_STREAM) {
        sections.push(parse_property_stream(
            "summary-information",
            SUMMARY_INFO_STREAM,
            &bytes,
        )?);
    }
    if let Some(bytes) = cfb.read_stream(DOC_SUMMARY_INFO_STREAM) {
        sections.push(parse_property_stream(
            "document-summary-information",
            DOC_SUMMARY_INFO_STREAM,
            &bytes,
        )?);
    }

    Ok(MetadataInspection {
        container: "cfb-ole".to_string(),
        section_count: sections.len(),
        sections,
    })
}

fn parse_property_stream(
    name: &str,
    path: &str,
    data: &[u8],
) -> Result<MetadataSection, ParserParseError> {
    if data.len() < 0x30 {
        return Err(ParserParseError::InvalidStructure(format!(
            "OLE property set stream {} is too short",
            path
        )));
    }
    let section_count = read_u32(data, 24)? as usize;
    const MAX_SECTIONS: usize = 1024;
    if section_count > MAX_SECTIONS {
        return Err(ParserParseError::InvalidStructure(format!(
            "OLE property set has too many sections ({section_count}, max {MAX_SECTIONS})"
        )));
    }
    if section_count == 0 {
        return Ok(MetadataSection {
            name: name.to_string(),
            path: path.to_string(),
            property_count: 0,
            properties: Vec::new(),
        });
    }

    let mut properties = Vec::new();
    for section_index in 0..section_count {
        let descriptor_offset = match 28usize.checked_add(section_index.saturating_mul(20)) {
            Some(off) => off,
            None => break,
        };
        if descriptor_offset + 20 > data.len() {
            break;
        }
        let section_offset = read_u32(data, descriptor_offset + 16)? as usize;
        if section_offset + 8 > data.len() {
            return Err(ParserParseError::InvalidStructure(format!(
                "OLE property set section offset is out of bounds for {}",
                path
            )));
        }

        let section_size = read_u32(data, section_offset)? as usize;
        let property_count = read_u32(data, section_offset + 4)? as usize;
        let section_end = section_offset.saturating_add(section_size).min(data.len());

        for index in 0..property_count {
            let entry_offset = match section_offset
                .checked_add(8)
                .and_then(|base| base.checked_add(index.saturating_mul(8)))
            {
                Some(off) => off,
                None => break,
            };
            if entry_offset + 8 > section_end {
                break;
            }
            let property_id = read_u32(data, entry_offset)?;
            let value_offset = read_u32(data, entry_offset + 4)? as usize;
            let absolute_offset = match section_offset.checked_add(value_offset) {
                Some(off) => off,
                None => continue,
            };
            if absolute_offset + 4 > section_end {
                continue;
            }
            let value_type = read_u32(data, absolute_offset)?;
            if let Some((type_name, value, display_value)) =
                parse_property_value(&data[absolute_offset..section_end], value_type)
            {
                properties.push(MetadataProperty {
                    id: property_id,
                    name: property_name(name, property_id).to_string(),
                    value_type: type_name.to_string(),
                    value,
                    display_value,
                });
            }
        }
    }

    Ok(MetadataSection {
        name: name.to_string(),
        path: path.to_string(),
        property_count: properties.len(),
        properties,
    })
}

fn parse_property_value(
    data: &[u8],
    value_type: u32,
) -> Option<(&'static str, String, Option<String>)> {
    match value_type {
        2 => Some(("i16", read_i16(data, 4).ok()?.to_string(), None)),
        3 => Some(("i32", read_i32(data, 4).ok()?.to_string(), None)),
        5 => Some(("f64", read_f64(data, 4).ok()?.to_string(), None)),
        11 => Some(("bool", (read_u16(data, 4).ok()? != 0).to_string(), None)),
        18 => Some(("u16", read_u16(data, 4).ok()?.to_string(), None)),
        19 => Some(("u32", read_u32(data, 4).ok()?.to_string(), None)),
        20 => Some(("i64", read_i64(data, 4).ok()?.to_string(), None)),
        30 => {
            let len = read_u32(data, 4).ok()? as usize;
            if 8 + len > data.len() {
                return None;
            }
            let bytes = &data[8..8 + len];
            let text = bytes.strip_suffix(&[0]).unwrap_or(bytes);
            Some(("lpstr", String::from_utf8_lossy(text).to_string(), None))
        }
        31 => {
            let chars = read_u32(data, 4).ok()? as usize;
            let byte_len = chars.checked_mul(2)?;
            let start = 8usize.checked_add(byte_len)?;
            if start > data.len() {
                return None;
            }
            let bytes = &data[8..start];
            let units: Vec<u16> = bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect();
            let text = String::from_utf16_lossy(units.strip_suffix(&[0]).unwrap_or(&units));
            Some(("lpwstr", text, None))
        }
        64 => {
            let raw = read_u64(data, 4).ok()?;
            Some(("filetime", raw.to_string(), Some(format_filetime_utc(raw))))
        }
        _ => None,
    }
}

fn format_filetime_utc(raw: u64) -> String {
    const WINDOWS_TO_UNIX_SECONDS: i128 = 11_644_473_600;
    const MAX_VALID_FILETIME: u64 = 600_000_000_000_000_000;
    if raw > MAX_VALID_FILETIME {
        return format!("filetime-overflow({})", raw);
    }
    let unix_seconds = (raw / 10_000_000) as i128 - WINDOWS_TO_UNIX_SECONDS;
    let days = unix_seconds.div_euclid(86_400);
    let secs_of_day = unix_seconds.rem_euclid(86_400);
    let (year, month, day) = civil_from_days(days);
    if !(1601..=9999).contains(&year) {
        return format!("filetime-invalid({})", raw);
    }
    let hour = secs_of_day / 3_600;
    let minute = (secs_of_day % 3_600) / 60;
    let second = secs_of_day % 60;
    format!("{year:04}-{month:02}-{day:02}T{hour:02}:{minute:02}:{second:02}Z")
}

fn civil_from_days(days: i128) -> (i128, i128, i128) {
    let z = days + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = mp + if mp < 10 { 3 } else { -9 };
    let year = y + if month <= 2 { 1 } else { 0 };
    (year, month, day)
}

fn property_name(section: &str, property_id: u32) -> &'static str {
    match section {
        "summary-information" => match property_id {
            1 => "codepage",
            2 => "title",
            3 => "subject",
            4 => "author",
            5 => "keywords",
            6 => "comments",
            7 => "template",
            8 => "last-saved-by",
            9 => "revision-number",
            10 => "edit-time",
            11 => "last-printed",
            12 => "created",
            13 => "modified",
            14 => "page-count",
            15 => "word-count",
            16 => "char-count",
            17 => "thumbnail",
            18 => "application-name",
            19 => "security",
            _ => "property",
        },
        "document-summary-information" => match property_id {
            1 => "codepage",
            2 => "category",
            3 => "presentation-format",
            4 => "byte-count",
            5 => "line-count",
            6 => "paragraph-count",
            7 => "slide-count",
            8 => "note-count",
            9 => "hidden-count",
            10 => "multimedia-count",
            11 => "scale",
            12 => "heading-pairs",
            13 => "titles-of-parts",
            14 => "manager",
            15 => "company",
            16 => "links-dirty",
            17 => "char-count-with-spaces",
            18 => "shared-document",
            19 => "link-base-updated",
            20 => "hyperlinks-changed",
            22 => "hyperlink-base",
            23 => "hlinks",
            24 => "mm-clips",
            26 => "content-type",
            27 => "content-status",
            28 => "language",
            29 => "document-version",
            _ => "property",
        },
        _ => "property",
    }
}

fn read_i16(data: &[u8], offset: usize) -> Result<i16, ParserParseError> {
    if offset + 2 > data.len() {
        return Err(ParserParseError::InvalidStructure(
            "OLE metadata read_i16 out of bounds".to_string(),
        ));
    }
    Ok(i16::from_le_bytes([data[offset], data[offset + 1]]))
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16, ParserParseError> {
    if offset + 2 > data.len() {
        return Err(ParserParseError::InvalidStructure(
            "OLE metadata read_u16 out of bounds".to_string(),
        ));
    }
    Ok(u16::from_le_bytes([data[offset], data[offset + 1]]))
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32, ParserParseError> {
    if offset + 4 > data.len() {
        return Err(ParserParseError::InvalidStructure(
            "OLE metadata read_u32 out of bounds".to_string(),
        ));
    }
    Ok(u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ]))
}

fn read_i32(data: &[u8], offset: usize) -> Result<i32, ParserParseError> {
    if offset + 4 > data.len() {
        return Err(ParserParseError::InvalidStructure(
            "OLE metadata read_i32 out of bounds".to_string(),
        ));
    }
    Ok(i32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ]))
}

fn read_u64(data: &[u8], offset: usize) -> Result<u64, ParserParseError> {
    if offset + 8 > data.len() {
        return Err(ParserParseError::InvalidStructure(
            "OLE metadata read_u64 out of bounds".to_string(),
        ));
    }
    Ok(u64::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
        data[offset + 4],
        data[offset + 5],
        data[offset + 6],
        data[offset + 7],
    ]))
}

fn read_i64(data: &[u8], offset: usize) -> Result<i64, ParserParseError> {
    if offset + 8 > data.len() {
        return Err(ParserParseError::InvalidStructure(
            "OLE metadata read_i64 out of bounds".to_string(),
        ));
    }
    Ok(i64::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
        data[offset + 4],
        data[offset + 5],
        data[offset + 6],
        data[offset + 7],
    ]))
}

fn read_f64(data: &[u8], offset: usize) -> Result<f64, ParserParseError> {
    if offset + 8 > data.len() {
        return Err(ParserParseError::InvalidStructure(
            "OLE metadata read_f64 out of bounds".to_string(),
        ));
    }
    Ok(f64::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
        data[offset + 4],
        data[offset + 5],
        data[offset + 6],
        data[offset + 7],
    ]))
}

#[cfg(test)]
mod tests {
    use super::{inspect_metadata_bytes, DOC_SUMMARY_INFO_STREAM, SUMMARY_INFO_STREAM};
    use crate::test_support::{build_test_cfb, build_test_property_set_stream, TestPropertyValue};

    #[test]
    fn inspect_metadata_reads_summary_and_doc_summary_streams() {
        let summary = build_test_property_set_stream(&[
            (2, TestPropertyValue::Str("Specimen")),
            (4, TestPropertyValue::Str("Analyst")),
            (12, TestPropertyValue::FileTime(100)),
            (14, TestPropertyValue::I32(7)),
        ]);
        let doc_summary = build_test_property_set_stream(&[
            (15, TestPropertyValue::WStr("ACME")),
            (16, TestPropertyValue::Bool(true)),
        ]);
        let inspection = inspect_metadata_bytes(&build_test_cfb(&[
            (SUMMARY_INFO_STREAM, &summary),
            (DOC_SUMMARY_INFO_STREAM, &doc_summary),
        ]))
        .expect("inspection");

        assert_eq!(inspection.section_count, 2);
        let summary = inspection
            .sections
            .iter()
            .find(|section| section.name == "summary-information")
            .expect("summary section");
        assert!(summary.properties.iter().any(|prop| prop.name == "title"
            && prop.value == "Specimen"
            && prop.display_value.is_none()));
        assert!(summary
            .properties
            .iter()
            .any(|prop| prop.name == "page-count" && prop.value == "7"));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "created"
                && prop.value == "100"
                && prop
                    .display_value
                    .as_deref()
                    .is_some_and(|value| value.ends_with('Z'))
        }));

        let doc_summary = inspection
            .sections
            .iter()
            .find(|section| section.name == "document-summary-information")
            .expect("doc summary section");
        assert!(doc_summary
            .properties
            .iter()
            .any(|prop| prop.name == "company" && prop.value == "ACME"));
        assert!(doc_summary
            .properties
            .iter()
            .any(|prop| prop.name == "links-dirty" && prop.value == "true"));
    }

    #[test]
    fn inspect_metadata_supports_additional_scalar_property_types() {
        let summary = build_test_property_set_stream(&[
            (1, TestPropertyValue::U16(1200)),
            (3, TestPropertyValue::Str("Specimen subject")),
            (4, TestPropertyValue::Str("Analyst")),
            (5, TestPropertyValue::Str("macro,ole")),
            (6, TestPropertyValue::Str("sample comment")),
            (7, TestPropertyValue::Str("Normal.dotm")),
            (8, TestPropertyValue::Str("Responder")),
            (9, TestPropertyValue::Str("7")),
            (10, TestPropertyValue::I64(3600)),
            (14, TestPropertyValue::I16(3)),
            (15, TestPropertyValue::U32(42)),
            (18, TestPropertyValue::Str("Microsoft Excel")),
            (19, TestPropertyValue::U32(1)),
        ]);
        let doc_summary = build_test_property_set_stream(&[
            (2, TestPropertyValue::Str("Malware triage")),
            (4, TestPropertyValue::I64(2048)),
            (14, TestPropertyValue::WStr("Analyst")),
            (15, TestPropertyValue::WStr("ACME")),
            (29, TestPropertyValue::WStr("16.0")),
            (12, TestPropertyValue::WStr("Slides")),
            (13, TestPropertyValue::WStr("Part A")),
            (26, TestPropertyValue::WStr("application/vnd.ms-excel")),
            (27, TestPropertyValue::WStr("final")),
            (23, TestPropertyValue::F64(2.5)),
            (28, TestPropertyValue::WStr("en-US")),
        ]);

        let inspection = inspect_metadata_bytes(&build_test_cfb(&[
            (SUMMARY_INFO_STREAM, &summary),
            (DOC_SUMMARY_INFO_STREAM, &doc_summary),
        ]))
        .expect("inspection");

        let summary = inspection
            .sections
            .iter()
            .find(|section| section.name == "summary-information")
            .expect("summary section");
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "codepage" && prop.value_type == "u16" && prop.value == "1200"
        }));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "subject" && prop.value_type == "lpstr" && prop.value == "Specimen subject"
        }));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "author" && prop.value_type == "lpstr" && prop.value == "Analyst"
        }));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "keywords" && prop.value_type == "lpstr" && prop.value == "macro,ole"
        }));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "comments" && prop.value_type == "lpstr" && prop.value == "sample comment"
        }));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "last-saved-by" && prop.value_type == "lpstr" && prop.value == "Responder"
        }));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "template" && prop.value_type == "lpstr" && prop.value == "Normal.dotm"
        }));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "revision-number" && prop.value_type == "lpstr" && prop.value == "7"
        }));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "edit-time" && prop.value_type == "i64" && prop.value == "3600"
        }));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "page-count" && prop.value_type == "i16" && prop.value == "3"
        }));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "word-count" && prop.value_type == "u32" && prop.value == "42"
        }));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "application-name"
                && prop.value_type == "lpstr"
                && prop.value == "Microsoft Excel"
        }));
        assert!(summary.properties.iter().any(|prop| {
            prop.name == "security" && prop.value_type == "u32" && prop.value == "1"
        }));

        let doc_summary = inspection
            .sections
            .iter()
            .find(|section| section.name == "document-summary-information")
            .expect("doc summary section");
        assert!(doc_summary.properties.iter().any(|prop| {
            prop.name == "category" && prop.value_type == "lpstr" && prop.value == "Malware triage"
        }));
        assert!(doc_summary.properties.iter().any(|prop| {
            prop.name == "byte-count" && prop.value_type == "i64" && prop.value == "2048"
        }));
        assert!(doc_summary.properties.iter().any(|prop| {
            prop.name == "heading-pairs" && prop.value_type == "lpwstr" && prop.value == "Slides"
        }));
        assert!(doc_summary.properties.iter().any(|prop| {
            prop.name == "titles-of-parts" && prop.value_type == "lpwstr" && prop.value == "Part A"
        }));
        assert!(doc_summary.properties.iter().any(|prop| {
            prop.name == "manager" && prop.value_type == "lpwstr" && prop.value == "Analyst"
        }));
        assert!(doc_summary.properties.iter().any(|prop| {
            prop.name == "company" && prop.value_type == "lpwstr" && prop.value == "ACME"
        }));
        assert!(doc_summary.properties.iter().any(|prop| {
            prop.name == "hlinks" && prop.value_type == "f64" && prop.value == "2.5"
        }));
        assert!(doc_summary.properties.iter().any(|prop| {
            prop.name == "content-type"
                && prop.value_type == "lpwstr"
                && prop.value == "application/vnd.ms-excel"
        }));
        assert!(doc_summary.properties.iter().any(|prop| {
            prop.name == "content-status" && prop.value_type == "lpwstr" && prop.value == "final"
        }));
        assert!(doc_summary.properties.iter().any(|prop| {
            prop.name == "language" && prop.value_type == "lpwstr" && prop.value == "en-US"
        }));
        assert!(doc_summary.properties.iter().any(|prop| {
            prop.name == "document-version" && prop.value_type == "lpwstr" && prop.value == "16.0"
        }));
    }
}
