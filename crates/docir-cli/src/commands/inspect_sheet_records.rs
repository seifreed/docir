//! Inspect low-level legacy XLS BIFF records.

use anyhow::Result;
use docir_app::{inspect_sheet_records_path, ParserConfig, SheetRecordInspection};
use std::path::PathBuf;

use crate::cli::JsonOutputOpts;
use crate::commands::util::{
    push_bullet_line, push_count_section, push_labeled_line, run_dual_output,
};

pub fn run(input: PathBuf, opts: JsonOutputOpts, parser_config: &ParserConfig) -> Result<()> {
    let JsonOutputOpts {
        json,
        pretty,
        output,
    } = opts;
    let inspection = inspect_sheet_records_path(&input, parser_config)?;
    run_dual_output(
        &inspection,
        "inspection",
        json,
        pretty,
        output,
        format_inspection_text,
    )
}

fn format_inspection_text(inspection: &SheetRecordInspection) -> String {
    let mut out = String::new();
    push_labeled_line(&mut out, 0, "Container", &inspection.container);
    push_labeled_line(&mut out, 0, "Workbook Stream", &inspection.workbook_stream);
    push_labeled_line(&mut out, 0, "Records", inspection.record_count);
    push_labeled_line(&mut out, 0, "Substreams", inspection.substream_count);
    push_labeled_line(&mut out, 0, "Anomalies", inspection.anomaly_count);
    push_count_section(
        &mut out,
        "Record Types",
        &inspection.record_type_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    push_count_section(
        &mut out,
        "Substreams",
        &inspection.substream_counts,
        |e| &e.bucket,
        |e| e.count,
    );
    if !inspection.records.is_empty() {
        out.push_str("\nRecords:\n");
        for record in &inspection.records {
            push_bullet_line(
                &mut out,
                2,
                &format!("0x{:04X}:{}", record.record_type, record.record_name),
                format!("offset={} size={}", record.offset, record.size),
            );
            push_labeled_line(&mut out, 4, "Substream Index", record.substream_index);
            push_labeled_line(&mut out, 4, "Substream Kind", &record.substream_kind);
        }
    }
    if !inspection.anomalies.is_empty() {
        out.push_str("\nAnomalies:\n");
        for anomaly in &inspection.anomalies {
            push_bullet_line(
                &mut out,
                2,
                &format!("{} @{}", anomaly.kind, anomaly.offset),
                &anomaly.message,
            );
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::{format_inspection_text, run};
    use crate::cli::JsonOutputOpts;
    use crate::test_support;
    use docir_app::{
        test_support::build_test_cfb, ParserConfig, SheetRecordAnomaly, SheetRecordCount,
        SheetRecordEntry, SheetRecordInspection,
    };
    use std::fs;

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
    fn inspect_sheet_records_run_writes_json() {
        let input = test_support::temp_file("legacy", "xls");
        let output = test_support::temp_file("legacy", "json");
        let mut workbook = Vec::new();
        workbook.extend(record(0x0809, &bof_payload(0x0005)));
        workbook.extend(record(0x0085, b"Sheet1"));
        workbook.extend(record(0x00FC, b"sst!"));
        workbook.extend(record(0x000A, &[]));
        workbook.extend(record(0x0809, &bof_payload(0x0010)));
        workbook.extend(record(0x0200, &[0; 10]));
        workbook.extend(record(0x0208, &[0; 16]));
        workbook.extend(record(0x00FD, &[0; 10]));
        workbook.extend(record(0x000A, &[]));
        fs::write(&input, build_test_cfb(&[("Workbook", &workbook)])).expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: true,
                pretty: true,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("inspect-sheet-records json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"workbook_stream\": \"Workbook\""));
        assert!(text.contains("\"record_type\": 2057"));
        assert!(text.contains("\"substream_kind\": \"worksheet\""));
        assert!(text.contains("\"record_name\": \"LABELSST\""));
    }

    #[test]
    fn inspect_sheet_records_run_writes_text() {
        let input = test_support::temp_file("legacy", "xls");
        let output = test_support::temp_file("legacy", "txt");
        let mut workbook = Vec::new();
        workbook.extend(record(0x0809, &bof_payload(0x0005)));
        workbook.extend(record(0x000A, &[]));
        fs::write(&input, build_test_cfb(&[("Workbook", &workbook)])).expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: false,
                pretty: false,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("inspect-sheet-records text");

        let text = fs::read_to_string(&output).expect("text output");
        assert!(text.contains("Workbook Stream: Workbook"));
        assert!(text.contains("0x0809:BOF"));
        assert!(text.contains("Substream Kind: workbook-globals"));
    }

    #[test]
    fn inspect_sheet_records_run_supports_book_stream_and_truncation() {
        let input = test_support::temp_file("legacy_book", "xls");
        let output = test_support::temp_file("legacy_book", "json");
        let mut workbook = Vec::new();
        workbook.extend(record(0x0809, &bof_payload(0x0020)));
        workbook.extend_from_slice(&0x0085u16.to_le_bytes());
        workbook.extend_from_slice(&8u16.to_le_bytes());
        workbook.extend_from_slice(b"abc");
        fs::write(&input, build_test_cfb(&[("Book", &workbook)])).expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: true,
                pretty: true,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("inspect-sheet-records book json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"workbook_stream\": \"Book\""));
        assert!(text.contains("\"substream_kind\": \"chart\""));
        assert!(text.contains("\"kind\": \"truncated-record\""));
    }

    #[test]
    fn inspect_sheet_records_run_handles_more_realistic_workbook_fixture() {
        let input = test_support::temp_file("legacy_realistic", "xls");
        let output = test_support::temp_file("legacy_realistic", "txt");
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
        workbook.extend(record(0x00FD, &[0; 10]));
        workbook.extend(record(0x023E, &[0; 18]));
        workbook.extend(record(0x001D, &[0; 9]));
        workbook.extend(record(0x000A, &[]));
        fs::write(&input, build_test_cfb(&[("Workbook", &workbook)])).expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: false,
                pretty: false,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("inspect-sheet-records realistic text");

        let text = fs::read_to_string(&output).expect("text output");
        assert!(text.contains("Substreams: 2"));
        assert!(text.contains("0x0042:CODEPAGE"));
        assert!(text.contains("0x0200:DIMENSIONS"));
        assert!(text.contains("0x023E:WINDOW2"));
        assert!(text.contains("0x001D:SELECTION"));
    }

    #[test]
    fn inspect_sheet_records_run_reports_dialog_sheet_substream() {
        let input = test_support::temp_file("legacy_dialog", "xls");
        let output = test_support::temp_file("legacy_dialog", "json");
        let mut workbook = Vec::new();
        workbook.extend(record(0x0809, &bof_payload(0x0005)));
        workbook.extend(record(0x0085, b"Dialog1"));
        workbook.extend(record(0x000A, &[]));
        workbook.extend(record(0x0809, &bof_payload(0x0200)));
        workbook.extend(record(0x0200, &[0; 10]));
        workbook.extend(record(0x023E, &[0; 18]));
        workbook.extend(record(0x000A, &[]));
        fs::write(&input, build_test_cfb(&[("Workbook", &workbook)])).expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: true,
                pretty: true,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("inspect-sheet-records dialog json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"substream_kind\": \"dialog-sheet\""));
        assert!(text.contains("\"record_name\": \"WINDOW2\""));
    }

    #[test]
    fn inspect_sheet_records_run_reports_macro_sheet_substream() {
        let input = test_support::temp_file("legacy_macro", "xls");
        let output = test_support::temp_file("legacy_macro", "txt");
        let mut workbook = Vec::new();
        workbook.extend(record(0x0809, &bof_payload(0x0005)));
        workbook.extend(record(0x0085, b"Macro1"));
        workbook.extend(record(0x000A, &[]));
        workbook.extend(record(0x0809, &bof_payload(0x0040)));
        workbook.extend(record(0x0200, &[0; 10]));
        workbook.extend(record(0x023E, &[0; 18]));
        workbook.extend(record(0x000A, &[]));
        fs::write(&input, build_test_cfb(&[("Workbook", &workbook)])).expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: false,
                pretty: false,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("inspect-sheet-records macro text");

        let text = fs::read_to_string(&output).expect("text output");
        assert!(text.contains("Substream Kind: macro-sheet"));
        assert!(text.contains("0x023E:WINDOW2"));
    }

    #[test]
    fn inspect_sheet_records_run_reports_chart_substream_in_json() {
        let input = test_support::temp_file("legacy_chart", "xls");
        let output = test_support::temp_file("legacy_chart", "json");
        let mut workbook = Vec::new();
        workbook.extend(record(0x0809, &bof_payload(0x0005)));
        workbook.extend(record(0x0085, b"Chart1"));
        workbook.extend(record(0x000A, &[]));
        workbook.extend(record(0x0809, &bof_payload(0x0020)));
        workbook.extend(record(0x023E, &[0; 18]));
        workbook.extend(record(0x000A, &[]));
        fs::write(&input, build_test_cfb(&[("Workbook", &workbook)])).expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: true,
                pretty: true,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("inspect-sheet-records chart json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"substream_kind\": \"chart\""));
        assert!(text.contains("\"record_name\": \"WINDOW2\""));
    }

    #[test]
    fn inspect_sheet_records_run_reports_workbook_globals_and_worksheet_json() {
        let input = test_support::temp_file("legacy_globals_ws", "xls");
        let output = test_support::temp_file("legacy_globals_ws", "json");
        let mut workbook = Vec::new();
        workbook.extend(record(0x0809, &bof_payload(0x0005)));
        workbook.extend(record(0x0042, &1200u16.to_le_bytes()));
        workbook.extend(record(0x0085, b"Sheet1"));
        workbook.extend(record(0x00FC, b"sst"));
        workbook.extend(record(0x000A, &[]));
        workbook.extend(record(0x0809, &bof_payload(0x0010)));
        workbook.extend(record(0x0200, &[0; 10]));
        workbook.extend(record(0x0208, &[0; 16]));
        workbook.extend(record(0x000A, &[]));
        fs::write(&input, build_test_cfb(&[("Workbook", &workbook)])).expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: true,
                pretty: true,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("inspect-sheet-records globals+worksheet json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"substream_kind\": \"workbook-globals\""));
        assert!(text.contains("\"substream_kind\": \"worksheet\""));
        assert!(text.contains("\"record_name\": \"CODEPAGE\""));
        assert!(text.contains("\"record_name\": \"DIMENSIONS\""));
    }

    #[test]
    fn inspect_sheet_records_run_reports_boundsheet_and_sst_in_text() {
        let input = test_support::temp_file("legacy_boundsheet_sst", "xls");
        let output = test_support::temp_file("legacy_boundsheet_sst", "txt");
        let mut workbook = Vec::new();
        workbook.extend(record(0x0809, &bof_payload(0x0005)));
        workbook.extend(record(0x0085, b"Sheet1"));
        workbook.extend(record(0x00FC, b"sst"));
        workbook.extend(record(0x000A, &[]));
        fs::write(&input, build_test_cfb(&[("Workbook", &workbook)])).expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: false,
                pretty: false,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("inspect-sheet-records boundsheet text");

        let text = fs::read_to_string(&output).expect("text output");
        assert!(text.contains("0x0085:BOUNDSHEET"));
        assert!(text.contains("0x00FC:SST"));
    }

    #[test]
    fn format_inspection_text_renders_expected_fields() {
        let inspection = SheetRecordInspection {
            container: "cfb".to_string(),
            workbook_stream: "Workbook".to_string(),
            record_count: 2,
            substream_count: 1,
            anomaly_count: 1,
            record_type_counts: vec![SheetRecordCount {
                bucket: "0x0809:BOF".to_string(),
                count: 1,
            }],
            substream_counts: vec![SheetRecordCount {
                bucket: "0:workbook-globals".to_string(),
                count: 2,
            }],
            records: vec![SheetRecordEntry {
                offset: 0,
                record_type: 0x0809,
                record_name: "BOF".to_string(),
                size: 4,
                substream_index: 0,
                substream_kind: "workbook-globals".to_string(),
            }],
            anomalies: vec![SheetRecordAnomaly {
                kind: "truncated-record".to_string(),
                offset: 12,
                message: "record 0x0085 declares 8 byte(s) but stream ends after 3".to_string(),
            }],
        };

        let text = format_inspection_text(&inspection);
        assert!(text.contains("Workbook Stream: Workbook"));
        assert!(text.contains("Record Types:"));
        assert!(text.contains("0x0809:BOF"));
        assert!(text.contains("Substream Kind: workbook-globals"));
        assert!(text.contains("truncated-record @12"));
    }
}
