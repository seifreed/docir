//! Low-level BIFF record reader for legacy XLS workbook streams.

use crate::error::ParseError;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsRecordScan {
    pub records: Vec<XlsRecordHeader>,
    pub anomalies: Vec<XlsRecordAnomaly>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsRecordHeader {
    pub offset: usize,
    pub record_type: u16,
    pub size: u16,
    pub record_name: &'static str,
    pub substream_index: usize,
    pub substream_kind: XlsSubstreamKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsRecordAnomaly {
    pub kind: &'static str,
    pub offset: usize,
    pub message: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsSubstreamKind {
    Unknown,
    WorkbookGlobals,
    Worksheet,
    Chart,
    MacroSheet,
    DialogSheet,
    VisualBasicModule,
    WorkspaceFile,
}

impl XlsSubstreamKind {
    /// Returns the stable display label for this BIFF substream kind.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Unknown => "unknown",
            Self::WorkbookGlobals => "workbook-globals",
            Self::Worksheet => "worksheet",
            Self::Chart => "chart",
            Self::MacroSheet => "macro-sheet",
            Self::DialogSheet => "dialog-sheet",
            Self::VisualBasicModule => "vb-module",
            Self::WorkspaceFile => "workspace-file",
        }
    }
}

/// Reads a legacy XLS workbook stream into BIFF record headers and anomalies.
pub fn read_xls_records(data: &[u8]) -> Result<XlsRecordScan, ParseError> {
    let mut offset = 0usize;
    let mut records = Vec::new();
    let mut anomalies = Vec::new();
    let mut substream_index = 0usize;
    let mut current_substream = XlsSubstreamKind::Unknown;

    while offset < data.len() {
        let remaining = data.len().saturating_sub(offset);
        if remaining < 4 {
            anomalies.push(XlsRecordAnomaly {
                kind: "trailing-bytes",
                offset,
                message: format!(
                    "{} trailing byte(s) after last complete record header",
                    remaining
                ),
            });
            break;
        }

        let record_type = u16::from_le_bytes([data[offset], data[offset + 1]]);
        let size = u16::from_le_bytes([data[offset + 2], data[offset + 3]]);
        let end = offset
            .checked_add(4)
            .and_then(|o| o.checked_add(size as usize));
        let end = match end {
            Some(e) if e <= data.len() => e,
            Some(_) => {
                anomalies.push(XlsRecordAnomaly {
                    kind: "truncated-record",
                    offset,
                    message: format!(
                        "record 0x{record_type:04X} declares {} byte(s) but stream ends after {}",
                        size,
                        data.len().saturating_sub(offset + 4)
                    ),
                });
                break;
            }
            None => {
                anomalies.push(XlsRecordAnomaly {
                    kind: "record-overflow",
                    offset,
                    message: format!("record 0x{record_type:04X} caused arithmetic overflow"),
                });
                break;
            }
        };

        let payload = &data[offset + 4..end];
        if record_type == 0x0809 {
            let next_kind = classify_bof_substream(payload);
            if !records.is_empty() {
                substream_index += 1;
            }
            current_substream = next_kind;
        }

        records.push(XlsRecordHeader {
            offset,
            record_type,
            size,
            record_name: record_name(record_type),
            substream_index,
            substream_kind: current_substream,
        });

        offset = end;
    }

    Ok(XlsRecordScan { records, anomalies })
}

fn classify_bof_substream(payload: &[u8]) -> XlsSubstreamKind {
    if payload.len() < 4 {
        return XlsSubstreamKind::Unknown;
    }
    let bof_type = u16::from_le_bytes([payload[2], payload[3]]);
    match bof_type {
        0x0005 => XlsSubstreamKind::WorkbookGlobals,
        0x0010 => XlsSubstreamKind::Worksheet,
        0x0020 => XlsSubstreamKind::Chart,
        0x0040 => XlsSubstreamKind::MacroSheet,
        0x0200 => XlsSubstreamKind::DialogSheet,
        0x0006 => XlsSubstreamKind::VisualBasicModule,
        0x0100 => XlsSubstreamKind::WorkspaceFile,
        _ => XlsSubstreamKind::Unknown,
    }
}

fn record_name(record_type: u16) -> &'static str {
    XLS_RECORD_NAMES
        .iter()
        .find_map(|(code, name)| (*code == record_type).then_some(*name))
        .unwrap_or("UNKNOWN")
}

const XLS_RECORD_NAMES: &[(u16, &str)] = &[
    (0x000A, "EOF"),
    (0x000C, "CALCCOUNT"),
    (0x000D, "CALCMODE"),
    (0x000F, "REFMODE"),
    (0x0010, "DELTA"),
    (0x0017, "EXTERNSHEET"),
    (0x0012, "PROTECT"),
    (0x0019, "WINDOWPROTECT"),
    (0x0022, "1904"),
    (0x0026, "LEFTMARGIN"),
    (0x0027, "RIGHTMARGIN"),
    (0x0028, "TOPMARGIN"),
    (0x0029, "BOTTOMMARGIN"),
    (0x002B, "PRINTROWCOL"),
    (0x002C, "PRINTGRIDLINES"),
    (0x002F, "FILEPASS"),
    (0x003C, "CONTINUE"),
    (0x003D, "WINDOW1"),
    (0x0042, "CODEPAGE"),
    (0x004D, "PLS"),
    (0x005A, "CRN"),
    (0x0055, "DEFCOLWIDTH"),
    (0x005C, "WRITEACCESS"),
    (0x005F, "SAVERECALC"),
    (0x0063, "OBJPROTECT"),
    (0x0078, "EXTERNNAME"),
    (0x007E, "RKREC"),
    (0x007D, "COLINFO"),
    (0x0080, "GUTS"),
    (0x0081, "WSBOOL"),
    (0x0085, "BOUNDSHEET"),
    (0x008C, "COUNTRY"),
    (0x008D, "HIDEOBJ"),
    (0x0090, "SORT"),
    (0x0092, "PALETTE"),
    (0x009C, "FNGROUPNAME"),
    (0x00A1, "SETUP"),
    (0x00A0, "SCENPROTECT"),
    (0x00AF, "FILESHARING"),
    (0x00BD, "MULRK"),
    (0x00BE, "MULBLANK"),
    (0x00E1, "INTERFACEHDR"),
    (0x00E2, "INTERFACEEND"),
    (0x00E3, "SXVS"),
    (0x00EB, "MSODRAWINGGROUP"),
    (0x00FC, "SST"),
    (0x0200, "DIMENSIONS"),
    (0x0201, "BLANK"),
    (0x0204, "LABEL"),
    (0x0206, "FORMULA"),
    (0x0203, "NUMBER"),
    (0x0208, "ROW"),
    (0x020B, "INDEX"),
    (0x0236, "TABLE"),
    (0x023E, "WINDOW2"),
    (0x027E, "RK"),
    (0x0291, "FONT"),
    (0x0294, "FORMAT"),
    (0x0293, "STYLE"),
    (0x00E0, "XF"),
    (0x00FD, "LABELSST"),
    (0x00E5, "MERGEDCELLS"),
    (0x001D, "SELECTION"),
    (0x041E, "FORMATLIST"),
    (0x0800, "HLINK"),
    (0x0809, "BOF"),
];

#[cfg(test)]
mod tests {
    use super::{read_xls_records, XlsSubstreamKind};

    fn record(record_type: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + payload.len());
        out.extend_from_slice(&record_type.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    fn bof_payload(substream_type: u16) -> [u8; 4] {
        let mut payload = [0u8; 4];
        payload[..2].copy_from_slice(&0x0600u16.to_le_bytes());
        payload[2..4].copy_from_slice(&substream_type.to_le_bytes());
        payload
    }

    #[test]
    fn read_xls_records_scans_multiple_substreams() {
        let mut stream = Vec::new();
        stream.extend(record(0x0809, &bof_payload(0x0005)));
        stream.extend(record(0x0085, b"sheet"));
        stream.extend(record(0x000A, &[]));
        stream.extend(record(0x0809, &bof_payload(0x0010)));
        stream.extend(record(0x0208, &[0; 16]));
        stream.extend(record(0x000A, &[]));

        let scan = read_xls_records(&stream).expect("scan");
        assert_eq!(scan.records.len(), 6);
        assert!(scan.anomalies.is_empty());
        assert_eq!(scan.records[0].record_name, "BOF");
        assert_eq!(
            scan.records[0].substream_kind,
            XlsSubstreamKind::WorkbookGlobals
        );
        assert_eq!(scan.records[3].substream_index, 1);
        assert_eq!(scan.records[3].substream_kind, XlsSubstreamKind::Worksheet);
    }

    #[test]
    fn read_xls_records_reports_truncated_record() {
        let mut stream = Vec::new();
        stream.extend(record(0x0809, &bof_payload(0x0005)));
        stream.extend_from_slice(&0x0085u16.to_le_bytes());
        stream.extend_from_slice(&8u16.to_le_bytes());
        stream.extend_from_slice(b"abc");

        let scan = read_xls_records(&stream).expect("scan");
        assert_eq!(scan.records.len(), 1);
        assert_eq!(scan.anomalies.len(), 1);
        assert_eq!(scan.anomalies[0].kind, "truncated-record");
    }

    #[test]
    fn read_xls_records_classifies_chart_and_dialog_substreams() {
        let mut stream = Vec::new();
        stream.extend(record(0x0809, &bof_payload(0x0020)));
        stream.extend(record(0x000A, &[]));
        stream.extend(record(0x0809, &bof_payload(0x0200)));
        stream.extend(record(0x000A, &[]));

        let scan = read_xls_records(&stream).expect("scan");
        assert_eq!(scan.records[0].substream_kind, XlsSubstreamKind::Chart);
        assert_eq!(
            scan.records[2].substream_kind,
            XlsSubstreamKind::DialogSheet
        );
    }
}
