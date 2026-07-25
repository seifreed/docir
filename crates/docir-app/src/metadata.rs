use crate::io_support::with_file_bytes;
use crate::ports::CfbStreamReaderPort;
use crate::{AppResult, ParserConfig, adapters};
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

type ParsedPropertyValue = (&'static str, String, Option<String>);

/// Inspect metadata from a legacy CFB/OLE file on disk.
pub fn inspect_metadata_path<P: AsRef<Path>>(
    path: P,
    config: &ParserConfig,
) -> AppResult<MetadataInspection> {
    with_file_bytes(path, config.max_input_size, inspect_metadata_bytes)
}

/// Inspect metadata from raw CFB/OLE bytes.
pub fn inspect_metadata_bytes(data: &[u8]) -> AppResult<MetadataInspection> {
    let reader = adapters::default_cfb_stream_reader();
    inspect_metadata_with_reader(data, &reader)
}

fn inspect_metadata_with_reader(
    data: &[u8],
    reader: &impl CfbStreamReaderPort,
) -> AppResult<MetadataInspection> {
    let mut sections = Vec::new();

    for (path, bytes) in
        reader.read_streams(data, &[SUMMARY_INFO_STREAM, DOC_SUMMARY_INFO_STREAM])?
    {
        let name = match path.as_str() {
            SUMMARY_INFO_STREAM => "summary-information",
            DOC_SUMMARY_INFO_STREAM => "document-summary-information",
            _ => continue,
        };
        sections.push(parse_property_stream(name, &path, &bytes)?);
    }

    Ok(MetadataInspection {
        container: "cfb-ole".to_string(),
        section_count: sections.len(),
        sections,
    })
}

fn parse_section_entries(
    name: &str,
    data: &[u8],
    section_index: usize,
    path: &str,
) -> Result<Vec<MetadataProperty>, ParserParseError> {
    const SECTION_HEADER_SIZE: usize = 8;
    const PROPERTY_ENTRY_SIZE: usize = 8;

    let descriptor_offset = match 28usize.checked_add(section_index.saturating_mul(20)) {
        Some(off) => off,
        None => return Ok(Vec::new()),
    };
    if descriptor_offset + 20 > data.len() {
        return Ok(Vec::new());
    }
    let section_offset = read_u32(data, descriptor_offset + 16)? as usize;
    if section_offset + SECTION_HEADER_SIZE > data.len() {
        return Err(ParserParseError::InvalidStructure(format!(
            "OLE property set section offset is out of bounds for {}",
            path
        )));
    }

    let section_size = read_u32(data, section_offset)? as usize;
    if section_size < SECTION_HEADER_SIZE {
        return Err(ParserParseError::InvalidStructure(
            "OLE property set section size is too small".to_string(),
        ));
    }
    let property_count = read_u32(data, section_offset + 4)? as usize;
    let section_end = section_offset.checked_add(section_size).ok_or_else(|| {
        ParserParseError::InvalidStructure("OLE property set section size overflow".to_string())
    })?;
    if section_end > data.len() {
        return Err(ParserParseError::InvalidStructure(
            "OLE property set section size exceeds stream length".to_string(),
        ));
    }
    let property_table_end = section_offset
        .checked_add(SECTION_HEADER_SIZE)
        .and_then(|base| base.checked_add(property_count.checked_mul(PROPERTY_ENTRY_SIZE)?))
        .ok_or_else(|| {
            ParserParseError::InvalidStructure(
                "OLE property set property table size overflow".to_string(),
            )
        })?;
    if property_table_end > section_end {
        return Err(ParserParseError::InvalidStructure(
            "OLE property set property table exceeds section size".to_string(),
        ));
    }

    let mut entries = Vec::new();
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
        let absolute_offset = section_offset.checked_add(value_offset).ok_or_else(|| {
            ParserParseError::InvalidStructure(
                "OLE property set property value offset overflow".to_string(),
            )
        })?;
        let value_type_end = absolute_offset.checked_add(4).ok_or_else(|| {
            ParserParseError::InvalidStructure(
                "OLE property set property value offset overflow".to_string(),
            )
        })?;
        if value_type_end > section_end {
            return Err(ParserParseError::InvalidStructure(
                "OLE property set property value offset exceeds section size".to_string(),
            ));
        }
        let value_type = read_u32(data, absolute_offset)?;
        if let Some((type_name, value, display_value)) =
            parse_property_value(&data[absolute_offset..section_end], value_type)?
        {
            entries.push(MetadataProperty {
                id: property_id,
                name: property_name(name, property_id).to_string(),
                value_type: type_name.to_string(),
                value,
                display_value,
            });
        }
    }
    Ok(entries)
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
        properties.extend(parse_section_entries(name, data, section_index, path)?);
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
) -> Result<Option<ParsedPropertyValue>, ParserParseError> {
    match value_type {
        2 => Ok(Some(("i16", read_i16(data, 4)?.to_string(), None))),
        3 => Ok(Some(("i32", read_i32(data, 4)?.to_string(), None))),
        5 => Ok(Some(("f64", read_f64(data, 4)?.to_string(), None))),
        11 => Ok(Some(("bool", (read_u16(data, 4)? != 0).to_string(), None))),
        18 => Ok(Some(("u16", read_u16(data, 4)?.to_string(), None))),
        19 => Ok(Some(("u32", read_u32(data, 4)?.to_string(), None))),
        20 => Ok(Some(("i64", read_i64(data, 4)?.to_string(), None))),
        30 => {
            let len = read_u32(data, 4)? as usize;
            if 8 + len > data.len() {
                return Err(ParserParseError::InvalidStructure(
                    "OLE metadata LPSTR value exceeds property bounds".to_string(),
                ));
            }
            let bytes = &data[8..8 + len];
            let text = bytes.strip_suffix(&[0]).unwrap_or(bytes);
            Ok(Some((
                "lpstr",
                String::from_utf8_lossy(text).to_string(),
                None,
            )))
        }
        31 => {
            let chars = read_u32(data, 4)? as usize;
            let byte_len = chars.checked_mul(2).ok_or_else(|| {
                ParserParseError::InvalidStructure(
                    "OLE metadata LPWSTR byte length overflow".to_string(),
                )
            })?;
            let start = 8usize.checked_add(byte_len).ok_or_else(|| {
                ParserParseError::InvalidStructure(
                    "OLE metadata LPWSTR value bounds overflow".to_string(),
                )
            })?;
            if start > data.len() {
                return Err(ParserParseError::InvalidStructure(
                    "OLE metadata LPWSTR value exceeds property bounds".to_string(),
                ));
            }
            let bytes = &data[8..start];
            let units: Vec<u16> = bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect();
            let text = String::from_utf16_lossy(units.strip_suffix(&[0]).unwrap_or(&units));
            Ok(Some(("lpwstr", text, None)))
        }
        64 => {
            let raw = read_u64(data, 4)?;
            Ok(Some((
                "filetime",
                raw.to_string(),
                Some(format_filetime_utc(raw)),
            )))
        }
        _ => Ok(None),
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
    Ok(i16::from_le_bytes(read_le_bytes(data, offset, "read_i16")?))
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16, ParserParseError> {
    Ok(u16::from_le_bytes(read_le_bytes(data, offset, "read_u16")?))
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32, ParserParseError> {
    Ok(u32::from_le_bytes(read_le_bytes(data, offset, "read_u32")?))
}

fn read_i32(data: &[u8], offset: usize) -> Result<i32, ParserParseError> {
    Ok(i32::from_le_bytes(read_le_bytes(data, offset, "read_i32")?))
}

fn read_u64(data: &[u8], offset: usize) -> Result<u64, ParserParseError> {
    Ok(u64::from_le_bytes(read_le_bytes(data, offset, "read_u64")?))
}

fn read_i64(data: &[u8], offset: usize) -> Result<i64, ParserParseError> {
    Ok(i64::from_le_bytes(read_le_bytes(data, offset, "read_i64")?))
}

fn read_f64(data: &[u8], offset: usize) -> Result<f64, ParserParseError> {
    Ok(f64::from_le_bytes(read_le_bytes(data, offset, "read_f64")?))
}

fn read_le_bytes<const N: usize>(
    data: &[u8],
    offset: usize,
    operation: &str,
) -> Result<[u8; N], ParserParseError> {
    let end = offset.checked_add(N).ok_or_else(|| {
        ParserParseError::InvalidStructure(format!("OLE metadata {operation} out of bounds"))
    })?;
    let bytes = data.get(offset..end).ok_or_else(|| {
        ParserParseError::InvalidStructure(format!("OLE metadata {operation} out of bounds"))
    })?;
    bytes.try_into().map_err(|_| {
        ParserParseError::InvalidStructure(format!("OLE metadata {operation} out of bounds"))
    })
}

#[cfg(test)]
#[path = "metadata_tests.rs"]
mod tests;
