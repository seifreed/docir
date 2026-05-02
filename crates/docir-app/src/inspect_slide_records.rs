use crate::io_support::read_bounded_file;
use crate::{AppResult, ParserConfig};
use docir_parser::ole::Cfb;
use docir_parser::ppt_records::{read_ppt_records, PptRecordScan};
use docir_parser::ParseError as ParserParseError;
use serde::Serialize;
use std::collections::BTreeMap;
use std::path::Path;

#[derive(Debug, Clone, Serialize)]
pub struct SlideRecordInspection {
    pub container: String,
    pub presentation_stream: String,
    pub record_count: usize,
    pub container_record_count: usize,
    pub max_depth: usize,
    pub anomaly_count: usize,
    pub record_type_counts: Vec<SlideRecordCount>,
    pub depth_counts: Vec<SlideRecordCount>,
    pub records: Vec<SlideRecordEntry>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub anomalies: Vec<SlideRecordAnomaly>,
}

#[derive(Debug, Clone, Serialize)]
pub struct SlideRecordCount {
    pub bucket: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct SlideRecordEntry {
    pub offset: usize,
    pub record_type: u16,
    pub record_name: String,
    pub version: u8,
    pub instance: u16,
    pub length: u32,
    pub is_container: bool,
    pub depth: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct SlideRecordAnomaly {
    pub kind: String,
    pub offset: usize,
    pub message: String,
}

pub fn inspect_slide_records_path<P: AsRef<Path>>(
    path: P,
    config: &ParserConfig,
) -> AppResult<SlideRecordInspection> {
    let bytes = read_bounded_file(path, config.max_input_size)?;
    inspect_slide_records_bytes(&bytes)
}

pub fn inspect_slide_records_bytes(data: &[u8]) -> AppResult<SlideRecordInspection> {
    let cfb = Cfb::parse(data.to_vec())?;
    let presentation_stream = "PowerPoint Document";
    let stream_bytes = cfb.read_stream(presentation_stream).ok_or_else(|| {
        ParserParseError::MissingPart(
            "legacy PPT presentation stream (PowerPoint Document)".to_string(),
        )
    })?;
    let scan = read_ppt_records(&stream_bytes)?;
    Ok(SlideRecordInspection::from_scan(presentation_stream, scan))
}

impl SlideRecordInspection {
    fn from_scan(presentation_stream: &str, scan: PptRecordScan) -> Self {
        let record_count = scan.records.len();
        let container_record_count = scan
            .records
            .iter()
            .filter(|record| record.is_container)
            .count();
        let max_depth = scan
            .records
            .iter()
            .map(|record| record.depth)
            .max()
            .unwrap_or(0);
        let anomaly_count = scan.anomalies.len();
        let record_type_counts = count_buckets(
            scan.records
                .iter()
                .map(|record| format!("0x{:04X}:{}", record.record_type, record.record_name)),
        );
        let depth_counts = count_buckets(
            scan.records
                .iter()
                .map(|record| format!("depth:{}", record.depth)),
        );
        let records = scan
            .records
            .into_iter()
            .map(|record| SlideRecordEntry {
                offset: record.offset,
                record_type: record.record_type,
                record_name: record.record_name.to_string(),
                version: record.version,
                instance: record.instance,
                length: record.length,
                is_container: record.is_container,
                depth: record.depth,
            })
            .collect();
        let anomalies = scan
            .anomalies
            .into_iter()
            .map(|anomaly| SlideRecordAnomaly {
                kind: anomaly.kind.to_string(),
                offset: anomaly.offset,
                message: anomaly.message,
            })
            .collect();

        Self {
            container: "cfb".to_string(),
            presentation_stream: presentation_stream.to_string(),
            record_count,
            container_record_count,
            max_depth,
            anomaly_count,
            record_type_counts,
            depth_counts,
            records,
            anomalies,
        }
    }
}

fn count_buckets(values: impl Iterator<Item = String>) -> Vec<SlideRecordCount> {
    let mut counts = BTreeMap::<String, usize>::new();
    for value in values {
        *counts.entry(value).or_default() += 1;
    }
    counts
        .into_iter()
        .map(|(bucket, count)| SlideRecordCount { bucket, count })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::inspect_slide_records_bytes;
    use crate::test_support::build_test_cfb;

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
    fn inspect_slide_records_reports_container_depth() {
        let child = record(0x00, 0, 0x03E9, &[1, 2, 3, 4]);
        let container = record(0x0F, 0, 0x03E8, &child);
        let bytes = build_test_cfb(&[
            ("PowerPoint Document", &container),
            ("Current User", b"user"),
        ]);
        let inspection = inspect_slide_records_bytes(&bytes).expect("inspect");
        assert_eq!(inspection.presentation_stream, "PowerPoint Document");
        assert_eq!(inspection.record_count, 2);
        assert_eq!(inspection.container_record_count, 1);
        assert_eq!(inspection.max_depth, 1);
    }

    #[test]
    fn inspect_slide_records_reports_deeper_nested_fixture() {
        let leaf = record(0x00, 0, 0x0409, &[1, 2]);
        let slide = record(0x0F, 0, 0x03F0, &leaf);
        let document = record(0x0F, 0, 0x03E8, &slide);
        let bytes = build_test_cfb(&[
            ("PowerPoint Document", &document),
            ("Current User", b"user"),
        ]);
        let inspection = inspect_slide_records_bytes(&bytes).expect("inspect");
        assert_eq!(inspection.record_count, 3);
        assert_eq!(inspection.container_record_count, 2);
        assert_eq!(inspection.max_depth, 2);
        assert!(inspection
            .depth_counts
            .iter()
            .any(|entry| entry.bucket == "depth:2"));
    }

    #[test]
    fn inspect_slide_records_reports_more_realistic_presentation_fixture() {
        let text_header = record(0x00, 0, 0x0409, &[1, 2]);
        let text_chars = record(0x00, 0, 0x03FA, &[0x41, 0x00, 0x42, 0x00]);
        let slide = record(0x0F, 0, 0x03F0, &[text_header, text_chars].concat());
        let env = record(0x0F, 0, 0x03FF, &record(0x00, 0, 0x03ED, &[0; 4]));
        let document = record(0x0F, 0, 0x03E8, &[env, slide].concat());
        let bytes = build_test_cfb(&[
            ("PowerPoint Document", &document),
            ("Current User", b"user"),
        ]);
        let inspection = inspect_slide_records_bytes(&bytes).expect("inspect");
        assert_eq!(inspection.max_depth, 2);
        assert!(inspection
            .record_type_counts
            .iter()
            .any(|entry| entry.bucket == "0x03FF:Environment"));
        assert!(inspection
            .record_type_counts
            .iter()
            .any(|entry| entry.bucket == "0x03FA:TextCharsAtom"));
    }

    #[test]
    fn inspect_slide_records_reports_truncated_record() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&u16::from(0x0Fu8).to_le_bytes());
        stream.extend_from_slice(&0x03E8u16.to_le_bytes());
        stream.extend_from_slice(&16u32.to_le_bytes());
        stream.extend_from_slice(b"abc");
        let bytes = build_test_cfb(&[("PowerPoint Document", &stream), ("Current User", b"user")]);
        let inspection = inspect_slide_records_bytes(&bytes).expect("inspect");
        assert_eq!(inspection.record_count, 0);
        assert_eq!(inspection.anomaly_count, 1);
        assert_eq!(inspection.anomalies[0].kind, "truncated-record");
    }
}
