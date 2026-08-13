//! Low-level record reader for legacy PowerPoint document streams.

use crate::error::ParseError;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptRecordScan {
    pub records: Vec<PptRecordHeader>,
    pub anomalies: Vec<PptRecordAnomaly>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptRecordHeader {
    pub offset: usize,
    pub record_type: u16,
    pub record_name: &'static str,
    pub version: u8,
    pub instance: u16,
    pub length: u32,
    pub is_container: bool,
    pub depth: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptRecordAnomaly {
    pub kind: &'static str,
    pub offset: usize,
    pub message: String,
}

/// Reads a legacy PowerPoint stream into flat record headers and anomalies.
pub fn read_ppt_records(data: &[u8]) -> Result<PptRecordScan, ParseError> {
    let mut records = Vec::new();
    let mut anomalies = Vec::new();
    scan_record_block(data, 0, 0, &mut records, &mut anomalies)?;

    Ok(PptRecordScan { records, anomalies })
}

const MAX_PPT_DEPTH: usize = 100;
const MAX_PPT_RECORDS: usize = 1_000_000;
const PPT_RECORD_HEADER_LEN: usize = 8;
const PPT_CONTAINER_VERSION: u8 = 0x0F;

struct PptRecordFrame {
    header: PptRecordHeader,
    payload_start: usize,
    end: usize,
}

fn scan_record_block(
    data: &[u8],
    block_start: usize,
    depth: usize,
    records: &mut Vec<PptRecordHeader>,
    anomalies: &mut Vec<PptRecordAnomaly>,
) -> Result<(), ParseError> {
    if depth > MAX_PPT_DEPTH {
        push_max_depth_anomaly(anomalies, block_start, depth);
        return Ok(());
    }

    let mut cursor = 0usize;
    while cursor < data.len() {
        let Some(frame) = read_record_frame(data, block_start, depth, cursor, anomalies) else {
            break;
        };

        let payload_start = frame.payload_start;
        let end = frame.end;
        let offset = frame.header.offset;
        let is_container = frame.header.is_container;
        let length = frame.header.length;
        if records.len() >= MAX_PPT_RECORDS {
            return Err(ParseError::ResourceLimit(format!(
                "PPT record count exceeds maximum ({MAX_PPT_RECORDS})"
            )));
        }
        records.push(frame.header);

        if is_container && length > 0 {
            scan_record_block(
                &data[payload_start..end],
                offset + PPT_RECORD_HEADER_LEN,
                depth + 1,
                records,
                anomalies,
            )?;
        }

        cursor = end;
    }
    Ok(())
}

fn read_record_frame(
    data: &[u8],
    block_start: usize,
    depth: usize,
    cursor: usize,
    anomalies: &mut Vec<PptRecordAnomaly>,
) -> Option<PptRecordFrame> {
    let offset = block_start + cursor;
    let remaining = data.len() - cursor;
    if remaining < PPT_RECORD_HEADER_LEN {
        push_trailing_bytes_anomaly(anomalies, offset, remaining);
        return None;
    }

    let ver_inst = u16::from_le_bytes([data[cursor], data[cursor + 1]]);
    let version = (ver_inst & 0x000F) as u8;
    let instance = ver_inst >> 4;
    let record_type = u16::from_le_bytes([data[cursor + 2], data[cursor + 3]]);
    let length = u32::from_le_bytes([
        data[cursor + 4],
        data[cursor + 5],
        data[cursor + 6],
        data[cursor + 7],
    ]);
    let payload_start = cursor.saturating_add(PPT_RECORD_HEADER_LEN);
    let end = payload_start.saturating_add(length as usize);
    if end > data.len() {
        push_truncated_record_anomaly(anomalies, data, cursor, offset, record_type, length);
        return None;
    }

    Some(PptRecordFrame {
        header: PptRecordHeader {
            offset,
            record_type,
            record_name: record_name(record_type),
            version,
            instance,
            length,
            is_container: version == PPT_CONTAINER_VERSION,
            depth,
        },
        payload_start,
        end,
    })
}

fn push_max_depth_anomaly(anomalies: &mut Vec<PptRecordAnomaly>, offset: usize, depth: usize) {
    anomalies.push(PptRecordAnomaly {
        kind: "max-depth-exceeded",
        offset,
        message: format!(
            "PPT record nesting depth {} exceeds maximum {}",
            depth, MAX_PPT_DEPTH
        ),
    });
}

fn push_trailing_bytes_anomaly(
    anomalies: &mut Vec<PptRecordAnomaly>,
    offset: usize,
    remaining: usize,
) {
    anomalies.push(PptRecordAnomaly {
        kind: "trailing-bytes",
        offset,
        message: format!(
            "{} trailing byte(s) after last complete PPT record header",
            remaining
        ),
    });
}

fn push_truncated_record_anomaly(
    anomalies: &mut Vec<PptRecordAnomaly>,
    data: &[u8],
    cursor: usize,
    offset: usize,
    record_type: u16,
    length: u32,
) {
    anomalies.push(PptRecordAnomaly {
        kind: "truncated-record",
        offset,
        message: format!(
            "record type 0x{record_type:04X} declares {} byte(s) but stream ends after {}",
            length,
            data.len().saturating_sub(cursor + PPT_RECORD_HEADER_LEN)
        ),
    });
}

fn record_name(record_type: u16) -> &'static str {
    match record_type {
        0x03E8 => "Document",
        0x03E9 => "DocumentAtom",
        0x03EA => "EndDocumentAtom",
        0x03EB => "SlidePersistDirectoryAtom",
        0x03EC => "NotesAtom",
        0x03ED => "EnvironmentAtom",
        0x03EE => "SlideBase",
        0x03EF => "SlideBaseAtom",
        0x03F3 => "SlideAtom",
        0x03F4 => "SlideListWithText",
        0x03F5 => "SlideListWithTextSubContainer",
        0x03F6 => "PPDrawing",
        0x03F7 => "ColorSchemeAtom",
        0x03F8 => "Notes",
        0x03F9 => "NotesAtomSecondary",
        0x03FA => "TextCharsAtom",
        0x03FB => "StyleTextPropAtom",
        0x03FC => "MasterTextPropAtom",
        0x03FD => "TextRulerAtom",
        0x03FE => "TextSpecInfoAtom",
        0x03FF => "Environment",
        0x0408 => "SlidePersistAtom",
        0x0409 => "TextHeaderAtom",
        0x0F00 => "ExObjList",
        0x0F01 => "ExObjListAtom",
        0x0FF5 => "ProgTags",
        0x0FF6 => "ProgBinaryTag",
        0x0F9F => "UserEditAtom",
        0x0FA0 => "CurrentUserAtom",
        0x0FA8 => "PersistPtrIncrementalBlock",
        0x1388 => "RoundTripContentMasterInfo12Atom",
        0x1772 => "RoundTripTheme12Atom",
        0x03F0 => "Slide",
        0x0400 => "Scheme",
        0x03F1 => "SlideAtomPlaceholder",
        _ => "UNKNOWN",
    }
}

#[cfg(test)]
mod tests {
    use super::{MAX_PPT_RECORDS, read_ppt_records};
    use crate::error::ParseError;

    fn record(version: u8, instance: u16, record_type: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(8 + payload.len());
        let ver_inst = ((instance << 4) | u16::from(version)).to_le_bytes();
        out.extend_from_slice(&ver_inst);
        out.extend_from_slice(&record_type.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    #[test]
    fn read_ppt_records_tracks_container_depth() {
        let child = record(0x00, 0, 0x03E9, &[1, 2, 3, 4]);
        let container = record(0x0F, 0, 0x03E8, &child);
        let scan = read_ppt_records(&container).expect("scan");
        assert_eq!(scan.records.len(), 2);
        assert!(scan.records[0].is_container);
        assert_eq!(scan.records[0].depth, 0);
        assert_eq!(scan.records[1].depth, 1);
        assert_eq!(scan.records[1].record_name, "DocumentAtom");
    }

    #[test]
    fn read_ppt_records_tracks_nested_containers() {
        let leaf = record(0x00, 0, 0x0409, &[1, 2]);
        let inner = record(0x0F, 0, 0x03F0, &leaf);
        let outer = record(0x0F, 0, 0x03E8, &inner);
        let scan = read_ppt_records(&outer).expect("scan");
        assert_eq!(scan.records.len(), 3);
        assert_eq!(scan.records[0].record_name, "Document");
        assert_eq!(scan.records[1].record_name, "Slide");
        assert_eq!(scan.records[1].depth, 1);
        assert_eq!(scan.records[2].record_name, "TextHeaderAtom");
        assert_eq!(scan.records[2].depth, 2);
    }

    #[test]
    fn read_ppt_records_rejects_excessive_record_count() {
        let mut stream = Vec::with_capacity((MAX_PPT_RECORDS + 1) * 8);
        for _ in 0..=MAX_PPT_RECORDS {
            stream.extend_from_slice(&record(0, 0, 0x0001, &[]));
        }

        let error = read_ppt_records(&stream).expect_err("record count must be bounded");
        assert!(
            matches!(error, ParseError::ResourceLimit(message) if message.contains("record count"))
        );
    }

    #[test]
    fn read_ppt_records_reports_truncated_record() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&((u16::from(0x0Fu8)).to_le_bytes()));
        stream.extend_from_slice(&0x03E8u16.to_le_bytes());
        stream.extend_from_slice(&16u32.to_le_bytes());
        stream.extend_from_slice(b"abc");

        let scan = read_ppt_records(&stream).expect("scan");
        assert!(scan.records.is_empty());
        assert_eq!(scan.anomalies.len(), 1);
        assert_eq!(scan.anomalies[0].kind, "truncated-record");
    }
}
