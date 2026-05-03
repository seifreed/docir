//! Inspect low-level legacy PPT binary records.

use anyhow::Result;
use docir_app::{inspect_slide_records_path, ParserConfig, SlideRecordInspection};
use std::path::PathBuf;

use crate::commands::util::{push_bullet_line, push_labeled_line, run_dual_output};

pub fn run(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let inspection = inspect_slide_records_path(&input, parser_config)?;
    run_dual_output(
        &inspection,
        "inspection",
        json,
        pretty,
        output,
        format_inspection_text,
    )
}

fn format_inspection_text(inspection: &SlideRecordInspection) -> String {
    let mut out = String::new();
    push_labeled_line(&mut out, 0, "Container", &inspection.container);
    push_labeled_line(
        &mut out,
        0,
        "Presentation Stream",
        &inspection.presentation_stream,
    );
    push_labeled_line(&mut out, 0, "Records", inspection.record_count);
    push_labeled_line(&mut out, 0, "Containers", inspection.container_record_count);
    push_labeled_line(&mut out, 0, "Max Depth", inspection.max_depth);
    push_labeled_line(&mut out, 0, "Anomalies", inspection.anomaly_count);
    if !inspection.record_type_counts.is_empty() {
        out.push_str("\nRecord Types:\n");
        for entry in &inspection.record_type_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.depth_counts.is_empty() {
        out.push_str("\nDepths:\n");
        for entry in &inspection.depth_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.records.is_empty() {
        out.push_str("\nRecords:\n");
        for record in &inspection.records {
            push_bullet_line(
                &mut out,
                2,
                &format!("0x{:04X}:{}", record.record_type, record.record_name),
                format!("offset={} len={}", record.offset, record.length),
            );
            push_labeled_line(&mut out, 4, "Version", record.version);
            push_labeled_line(&mut out, 4, "Instance", record.instance);
            push_labeled_line(&mut out, 4, "Container", record.is_container);
            push_labeled_line(&mut out, 4, "Depth", record.depth);
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
    use crate::test_support;
    use docir_app::{
        test_support::build_test_cfb, ParserConfig, SlideRecordAnomaly, SlideRecordCount,
        SlideRecordEntry, SlideRecordInspection,
    };
    use std::fs;

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
    fn inspect_slide_records_run_writes_json() {
        let input = test_support::temp_file("legacy", "ppt");
        let output = test_support::temp_file("legacy", "json");
        let leaf = record(0x00, 0, 0x0409, &[1, 2]);
        let slide = record(0x0F, 0, 0x03F0, &leaf);
        let container = record(0x0F, 0, 0x03E8, &slide);
        fs::write(
            &input,
            build_test_cfb(&[
                ("PowerPoint Document", &container),
                ("Current User", b"user"),
            ]),
        )
        .expect("fixture");

        run(
            input.clone(),
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inspect-slide-records json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"presentation_stream\": \"PowerPoint Document\""));
        assert!(text.contains("\"record_name\": \"Document\""));
        assert!(text.contains("\"record_name\": \"Slide\""));
        assert!(text.contains("\"depth\": 2"));
    }

    #[test]
    fn inspect_slide_records_run_writes_text() {
        let input = test_support::temp_file("legacy", "ppt");
        let output = test_support::temp_file("legacy", "txt");
        let atom = record(0x00, 0, 0x03E9, &[1, 2, 3, 4]);
        fs::write(
            &input,
            build_test_cfb(&[("PowerPoint Document", &atom), ("Current User", b"user")]),
        )
        .expect("fixture");

        run(
            input.clone(),
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inspect-slide-records text");

        let text = fs::read_to_string(&output).expect("text output");
        assert!(text.contains("Presentation Stream: PowerPoint Document"));
        assert!(text.contains("0x03E9:DocumentAtom"));
    }

    #[test]
    fn inspect_slide_records_run_handles_more_realistic_nested_fixture() {
        let input = test_support::temp_file("legacy_nested", "ppt");
        let output = test_support::temp_file("legacy_nested", "txt");
        let text_header = record(0x00, 0, 0x0409, &[1, 2]);
        let text_chars = record(0x00, 0, 0x03FA, &[0x41, 0x00, 0x42, 0x00]);
        let text_props = record(0x00, 0, 0x03FB, &[0; 4]);
        let slide = record(
            0x0F,
            0,
            0x03F0,
            &[text_header, text_chars, text_props].concat(),
        );
        let env = record(0x0F, 0, 0x03FF, &record(0x00, 0, 0x03ED, &[0; 4]));
        let ex_obj = record(0x0F, 0, 0x0F00, &record(0x00, 0, 0x0F01, &[0; 4]));
        let document = record(0x0F, 0, 0x03E8, &[env, ex_obj, slide].concat());
        fs::write(
            &input,
            build_test_cfb(&[
                ("PowerPoint Document", &document),
                ("Current User", b"user"),
            ]),
        )
        .expect("fixture");

        run(
            input.clone(),
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inspect-slide-records nested text");

        let text = fs::read_to_string(&output).expect("text output");
        assert!(text.contains("Max Depth: 2"));
        assert!(text.contains("0x03FF:Environment"));
        assert!(text.contains("0x0F00:ExObjList"));
        assert!(text.contains("0x0F01:ExObjListAtom"));
        assert!(text.contains("0x03FA:TextCharsAtom"));
        assert!(text.contains("0x03FB:StyleTextPropAtom"));
    }

    #[test]
    fn inspect_slide_records_run_reports_prog_tags_fixture() {
        let input = test_support::temp_file("legacy_prog_tags", "ppt");
        let output = test_support::temp_file("legacy_prog_tags", "json");
        let prog_binary = record(0x00, 0, 0x0FF6, &[0; 4]);
        let prog_tags = record(0x0F, 0, 0x0FF5, &prog_binary);
        let doc = record(0x0F, 0, 0x03E8, &prog_tags);
        fs::write(
            &input,
            build_test_cfb(&[("PowerPoint Document", &doc), ("Current User", b"user")]),
        )
        .expect("fixture");

        run(
            input.clone(),
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inspect-slide-records prog-tags json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"record_name\": \"ProgTags\""));
        assert!(text.contains("\"record_name\": \"ProgBinaryTag\""));
    }

    #[test]
    fn inspect_slide_records_run_reports_roundtrip_theme_fixture() {
        let input = test_support::temp_file("legacy_roundtrip", "ppt");
        let output = test_support::temp_file("legacy_roundtrip", "txt");
        let theme = record(0x00, 0, 0x1772, &[0; 4]);
        let slide = record(0x0F, 0, 0x03F0, &theme);
        let doc = record(0x0F, 0, 0x03E8, &slide);
        fs::write(
            &input,
            build_test_cfb(&[("PowerPoint Document", &doc), ("Current User", b"user")]),
        )
        .expect("fixture");

        run(
            input.clone(),
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inspect-slide-records roundtrip text");

        let text = fs::read_to_string(&output).expect("text output");
        assert!(text.contains("0x1772:RoundTripTheme12Atom"));
        assert!(text.contains("Max Depth: 2"));
    }

    #[test]
    fn inspect_slide_records_run_reports_user_edit_atom_in_json() {
        let input = test_support::temp_file("legacy_user_edit", "ppt");
        let output = test_support::temp_file("legacy_user_edit", "json");
        let user_edit = record(0x00, 0, 0x0F9F, &[0; 4]);
        let doc = record(0x0F, 0, 0x03E8, &user_edit);
        fs::write(
            &input,
            build_test_cfb(&[("PowerPoint Document", &doc), ("Current User", b"user")]),
        )
        .expect("fixture");

        run(
            input.clone(),
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inspect-slide-records user-edit json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"record_name\": \"UserEditAtom\""));
    }

    #[test]
    fn inspect_slide_records_run_reports_environment_and_slide_json() {
        let input = test_support::temp_file("legacy_env_slide", "ppt");
        let output = test_support::temp_file("legacy_env_slide", "json");
        let env = record(0x0F, 0, 0x03FF, &record(0x00, 0, 0x03ED, &[0; 4]));
        let slide = record(0x0F, 0, 0x03F0, &record(0x00, 0, 0x03FA, &[0x41, 0x00]));
        let doc = record(0x0F, 0, 0x03E8, &[env, slide].concat());
        fs::write(
            &input,
            build_test_cfb(&[("PowerPoint Document", &doc), ("Current User", b"user")]),
        )
        .expect("fixture");

        run(
            input.clone(),
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inspect-slide-records env slide json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"record_name\": \"Environment\""));
        assert!(text.contains("\"record_name\": \"Slide\""));
        assert!(text.contains("\"depth\": 2"));
    }

    #[test]
    fn inspect_slide_records_run_reports_document_atom_and_notes_text() {
        let input = test_support::temp_file("legacy_notes", "ppt");
        let output = test_support::temp_file("legacy_notes", "txt");
        let notes = record(0x00, 0, 0x03EC, &[0; 4]);
        let atom = record(0x00, 0, 0x03E9, &[0; 4]);
        let doc = record(0x0F, 0, 0x03E8, &[atom, notes].concat());
        fs::write(
            &input,
            build_test_cfb(&[("PowerPoint Document", &doc), ("Current User", b"user")]),
        )
        .expect("fixture");

        run(
            input.clone(),
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inspect-slide-records notes text");

        let text = fs::read_to_string(&output).expect("text output");
        assert!(text.contains("0x03E9:DocumentAtom"));
        assert!(text.contains("0x03EC:NotesAtom"));
    }

    #[test]
    fn format_inspection_text_renders_expected_fields() {
        let inspection = SlideRecordInspection {
            container: "cfb".to_string(),
            presentation_stream: "PowerPoint Document".to_string(),
            record_count: 2,
            container_record_count: 1,
            max_depth: 1,
            anomaly_count: 1,
            record_type_counts: vec![SlideRecordCount {
                bucket: "0x03E8:Document".to_string(),
                count: 1,
            }],
            depth_counts: vec![SlideRecordCount {
                bucket: "depth:1".to_string(),
                count: 1,
            }],
            records: vec![SlideRecordEntry {
                offset: 0,
                record_type: 0x03E8,
                record_name: "Document".to_string(),
                version: 0x0F,
                instance: 0,
                length: 12,
                is_container: true,
                depth: 0,
            }],
            anomalies: vec![SlideRecordAnomaly {
                kind: "truncated-record".to_string(),
                offset: 20,
                message: "record type 0x03E9 declares 16 byte(s) but stream ends after 3"
                    .to_string(),
            }],
        };

        let text = format_inspection_text(&inspection);
        assert!(text.contains("Presentation Stream: PowerPoint Document"));
        assert!(text.contains("Containers: 1"));
        assert!(text.contains("0x03E8:Document"));
        assert!(text.contains("truncated-record @20"));
    }
}
