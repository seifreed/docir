use crate::io_support::with_file_bytes;
use crate::{AppResult, ParserConfig};
use docir_parser::ole::Cfb;
use docir_parser::xls_records::{read_xls_records, XlsRecordScan, XlsSubstreamKind};
use docir_parser::ParseError as ParserParseError;
use serde::Serialize;
use std::collections::BTreeMap;
use std::path::Path;

#[derive(Debug, Clone, Serialize)]
pub struct SheetRecordInspection {
    pub container: String,
    pub workbook_stream: String,
    pub record_count: usize,
    pub substream_count: usize,
    pub anomaly_count: usize,
    pub record_type_counts: Vec<SheetRecordCount>,
    pub substream_counts: Vec<SheetRecordCount>,
    pub records: Vec<SheetRecordEntry>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub anomalies: Vec<SheetRecordAnomaly>,
}

#[derive(Debug, Clone, Serialize)]
pub struct SheetRecordCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct SheetRecordEntry {
    pub offset: usize,
    pub record_type: u16,
    pub record_name: String,
    pub size: u16,
    pub substream_index: usize,
    pub substream_kind: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct SheetRecordAnomaly {
    pub kind: String,
    pub offset: usize,
    pub message: String,
}

/// Inspects low-level BIFF records from a legacy XLS file path.
pub fn inspect_sheet_records_path<P: AsRef<Path>>(
    path: P,
    config: &ParserConfig,
) -> AppResult<SheetRecordInspection> {
    with_file_bytes(path, config.max_input_size, inspect_sheet_records_bytes)
}

/// Inspects low-level BIFF records from raw CFB/OLE bytes.
pub fn inspect_sheet_records_bytes(data: &[u8]) -> AppResult<SheetRecordInspection> {
    let cfb = Cfb::parse(data.to_vec())?;
    let workbook_stream = if cfb.has_stream("Workbook") {
        "Workbook"
    } else if cfb.has_stream("Book") {
        "Book"
    } else {
        return Err(ParserParseError::MissingPart(
            "legacy XLS workbook stream (Workbook/Book)".to_string(),
        )
        .into());
    };

    let workbook_bytes = cfb.read_stream(workbook_stream).ok_or_else(|| {
        ParserParseError::InvalidStructure(format!(
            "failed to read workbook stream {workbook_stream}"
        ))
    })?;
    let scan = read_xls_records(&workbook_bytes)?;
    Ok(SheetRecordInspection::from_scan(workbook_stream, scan))
}

impl SheetRecordInspection {
    fn from_scan(workbook_stream: &str, scan: XlsRecordScan) -> Self {
        let record_count = scan.records.len();
        let substream_count = scan
            .records
            .iter()
            .map(|record| record.substream_index)
            .max()
            .map(|value| value + 1)
            .unwrap_or(0);
        let anomaly_count = scan.anomalies.len();
        let record_type_counts = count_buckets(
            scan.records
                .iter()
                .map(|record| format!("0x{:04X}:{}", record.record_type, record.record_name)),
        );
        let substream_counts = count_buckets(scan.records.iter().map(|record| {
            format!(
                "{}:{}",
                record.substream_index,
                record.substream_kind.as_str()
            )
        }));
        let records = scan
            .records
            .into_iter()
            .map(|record| SheetRecordEntry {
                offset: record.offset,
                record_type: record.record_type,
                record_name: record.record_name.to_string(),
                size: record.size,
                substream_index: record.substream_index,
                substream_kind: substream_kind_label(record.substream_kind).to_string(),
            })
            .collect();
        let anomalies = scan
            .anomalies
            .into_iter()
            .map(|anomaly| SheetRecordAnomaly {
                kind: anomaly.kind.to_string(),
                offset: anomaly.offset,
                message: anomaly.message,
            })
            .collect();

        Self {
            container: "cfb".to_string(),
            workbook_stream: workbook_stream.to_string(),
            record_count,
            substream_count,
            anomaly_count,
            record_type_counts,
            substream_counts,
            records,
            anomalies,
        }
    }
}

fn substream_kind_label(kind: XlsSubstreamKind) -> &'static str {
    kind.as_str()
}

fn count_buckets(values: impl Iterator<Item = String>) -> Vec<SheetRecordCount> {
    let mut counts = BTreeMap::<String, usize>::new();
    for value in values {
        *counts.entry(value).or_default() += 1;
    }
    counts
        .into_iter()
        .map(|(bucket, count)| SheetRecordCount { bucket, count })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::{inspect_sheet_records_bytes, substream_kind_label};
    use crate::test_support::build_test_cfb;
    use docir_parser::xls_records::XlsSubstreamKind;

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
    fn inspect_sheet_records_reports_workbook_and_worksheet_substreams() {
        let mut workbook = Vec::new();
        workbook.extend(record(0x0809, &bof_payload(0x0005)));
        workbook.extend(record(0x0085, b"Sheet1"));
        workbook.extend(record(0x000A, &[]));
        workbook.extend(record(0x0809, &bof_payload(0x0010)));
        workbook.extend(record(0x0208, &[0; 16]));
        workbook.extend(record(0x000A, &[]));

        let bytes = build_test_cfb(&[("Workbook", &workbook)]);
        let inspection = inspect_sheet_records_bytes(&bytes).expect("inspect");
        assert_eq!(inspection.workbook_stream, "Workbook");
        assert_eq!(inspection.record_count, 6);
        assert_eq!(inspection.substream_count, 2);
        assert!(inspection
            .substream_counts
            .iter()
            .any(|entry| entry.bucket == "0:workbook-globals"));
        assert!(inspection
            .substream_counts
            .iter()
            .any(|entry| entry.bucket == "1:worksheet"));
    }

    #[test]
    fn inspect_sheet_records_reports_truncated_stream_anomaly() {
        let mut workbook = Vec::new();
        workbook.extend(record(0x0809, &bof_payload(0x0005)));
        workbook.extend_from_slice(&0x0085u16.to_le_bytes());
        workbook.extend_from_slice(&8u16.to_le_bytes());
        workbook.extend_from_slice(b"abc");

        let bytes = build_test_cfb(&[("Workbook", &workbook)]);
        let inspection = inspect_sheet_records_bytes(&bytes).expect("inspect");
        assert_eq!(inspection.record_count, 1);
        assert_eq!(inspection.anomaly_count, 1);
        assert_eq!(inspection.anomalies[0].kind, "truncated-record");
    }

    #[test]
    fn inspect_sheet_records_supports_book_stream_name() {
        let mut workbook = Vec::new();
        workbook.extend(record(0x0809, &bof_payload(0x0020)));
        workbook.extend(record(0x000A, &[]));

        let bytes = build_test_cfb(&[("Book", &workbook)]);
        let inspection = inspect_sheet_records_bytes(&bytes).expect("inspect");
        assert_eq!(inspection.workbook_stream, "Book");
        assert!(inspection
            .substream_counts
            .iter()
            .any(|entry| entry.bucket == "0:chart"));
    }

    #[test]
    fn inspect_sheet_records_reports_more_realistic_workbook_fixture() {
        let mut workbook = Vec::new();
        workbook.extend(record(0x0809, &bof_payload(0x0005)));
        workbook.extend(record(0x0042, &1200u16.to_le_bytes()));
        workbook.extend(record(0x0085, b"Sheet1"));
        workbook.extend(record(0x00FC, b"sst-data"));
        workbook.extend(record(0x00E0, &[0; 20]));
        workbook.extend(record(0x000A, &[]));
        workbook.extend(record(0x0809, &bof_payload(0x0010)));
        workbook.extend(record(0x0200, &[0; 10]));
        workbook.extend(record(0x0208, &[0; 16]));
        workbook.extend(record(0x00FD, &[0; 12]));
        workbook.extend(record(0x023E, &[0; 18]));
        workbook.extend(record(0x000A, &[]));

        let bytes = build_test_cfb(&[("Workbook", &workbook)]);
        let inspection = inspect_sheet_records_bytes(&bytes).expect("inspect");
        assert_eq!(inspection.substream_count, 2);
        assert!(inspection
            .record_type_counts
            .iter()
            .any(|entry| entry.bucket == "0x0042:CODEPAGE"));
        assert!(inspection
            .record_type_counts
            .iter()
            .any(|entry| entry.bucket == "0x023E:WINDOW2"));
    }

    #[test]
    fn substream_kind_label_uses_stable_values() {
        assert_eq!(
            substream_kind_label(XlsSubstreamKind::WorkbookGlobals),
            "workbook-globals"
        );
        assert_eq!(
            substream_kind_label(XlsSubstreamKind::Worksheet),
            "worksheet"
        );
    }
}
