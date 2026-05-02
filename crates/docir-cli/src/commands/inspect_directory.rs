//! Inspect normal CFB directory entries.

use anyhow::Result;
use docir_app::{inspect_directory_path, DirectoryInspection, ParserConfig};
use serde::Serialize;
use std::path::PathBuf;

use crate::commands::util::{
    push_bullet_line, push_labeled_line, write_json_output, write_text_output,
};

#[derive(Debug, Serialize)]
struct InspectDirectoryResult {
    inspection: DirectoryInspection,
}

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let inspection = inspect_directory_path(&input, parser_config)?;

    if json {
        return write_json_output(&InspectDirectoryResult { inspection }, pretty, output);
    }

    let text = format_inspection_text(&inspection);
    write_text_output(&text, output)
}

fn format_inspection_text(inspection: &DirectoryInspection) -> String {
    let mut out = String::new();
    push_labeled_line(&mut out, 0, "Container", &inspection.container);
    push_labeled_line(&mut out, 0, "Entries", inspection.entry_count);
    push_labeled_line(&mut out, 0, "Directory Score", &inspection.directory_score);
    if !inspection.role_counts.is_empty() {
        out.push_str("\nRole Counts:\n");
        for entry in &inspection.role_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.anomaly_counts.is_empty() {
        out.push_str("\nAnomalies:\n");
        for entry in &inspection.anomaly_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.anomaly_catalog.is_empty() {
        out.push_str("\nAnomaly Severity Catalog:\n");
        for entry in &inspection.anomaly_catalog {
            push_bullet_line(&mut out, 2, &entry.anomaly, &entry.severity);
        }
    }
    if !inspection.anomaly_severity_counts.is_empty() {
        out.push_str("\nAnomaly Severity Summary:\n");
        for entry in &inspection.anomaly_severity_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.reference_counts.is_empty() {
        out.push_str("\nReference Summary:\n");
        for entry in &inspection.reference_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.pointer_counts.is_empty() {
        out.push_str("\nPointer Summary:\n");
        for entry in &inspection.pointer_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.tree_density_counts.is_empty() {
        out.push_str("\nTree Density Summary:\n");
        for entry in &inspection.tree_density_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.dangling_state_counts.is_empty() {
        out.push_str("\nDangling By State:\n");
        for entry in &inspection.dangling_state_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.self_reference_counts.is_empty() {
        out.push_str("\nSelf References:\n");
        for entry in &inspection.self_reference_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.short_cycle_counts.is_empty() {
        out.push_str("\nShort Cycles:\n");
        for entry in &inspection.short_cycle_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.reachability_counts.is_empty() {
        out.push_str("\nReachability Summary:\n");
        for entry in &inspection.reachability_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.incoming_source_counts.is_empty() {
        out.push_str("\nIncoming Source Summary:\n");
        for entry in &inspection.incoming_source_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.incoming_source_type_counts.is_empty() {
        out.push_str("\nIncoming Source Types:\n");
        for entry in &inspection.incoming_source_type_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.dead_reference_counts.is_empty() {
        out.push_str("\nDead But Referenced:\n");
        for entry in &inspection.dead_reference_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.fanout_counts.is_empty() {
        out.push_str("\nFanout Summary:\n");
        for entry in &inspection.fanout_counts {
            push_bullet_line(&mut out, 2, &entry.bucket, entry.count);
        }
    }
    if !inspection.entries.is_empty() {
        out.push_str("\nDirectory:\n");
        for entry in &inspection.entries {
            push_bullet_line(&mut out, 2, &entry.entry_type, &entry.path);
            push_labeled_line(&mut out, 4, "Entry Index", entry.entry_index);
            push_labeled_line(&mut out, 4, "Name Length Raw", entry.name_len_raw);
            push_labeled_line(&mut out, 4, "Object Type Raw", entry.object_type_raw);
            push_labeled_line(&mut out, 4, "Color Flag Raw", entry.color_flag_raw);
            push_labeled_line(&mut out, 4, "State", &entry.state);
            push_labeled_line(&mut out, 4, "Classification", &entry.classification);
            push_labeled_line(&mut out, 4, "Anomaly Severity", &entry.anomaly_severity);
            if !entry.anomaly_tags.is_empty() {
                push_labeled_line(&mut out, 4, "Anomaly Tags", entry.anomaly_tags.join(", "));
            }
            if !entry.short_cycles.is_empty() {
                push_labeled_line(&mut out, 4, "Short Cycles", entry.short_cycles.join(", "));
            }
            push_labeled_line(
                &mut out,
                4,
                "Reachable From Root",
                entry.reachable_from_root,
            );
            push_labeled_line(
                &mut out,
                4,
                "Incoming References",
                entry.incoming_reference_count,
            );
            push_labeled_line(
                &mut out,
                4,
                "Incoming From Normal",
                entry.incoming_normal_reference_count,
            );
            push_labeled_line(
                &mut out,
                4,
                "Incoming From Anomalous",
                entry.incoming_anomalous_reference_count,
            );
            push_labeled_line(
                &mut out,
                4,
                "Incoming From Root Storage",
                entry.incoming_from_root_storage_count,
            );
            push_labeled_line(
                &mut out,
                4,
                "Incoming From Storage",
                entry.incoming_from_storage_count,
            );
            push_labeled_line(
                &mut out,
                4,
                "Incoming From Stream",
                entry.incoming_from_stream_count,
            );
            push_labeled_line(&mut out, 4, "Fanout", entry.fanout_count);
            if !entry.incoming_from.is_empty() {
                push_labeled_line(&mut out, 4, "Incoming From", entry.incoming_from.join(", "));
            }
            push_labeled_line(&mut out, 4, "Size", format!("{} bytes", entry.size_bytes));
            push_labeled_line(&mut out, 4, "Sector", entry.start_sector);
            push_labeled_line(&mut out, 4, "Left Sibling Raw", entry.left_sibling_raw);
            push_labeled_line(&mut out, 4, "Right Sibling Raw", entry.right_sibling_raw);
            push_labeled_line(&mut out, 4, "Child Raw", entry.child_raw);
            if let Some(left) = entry.left_sibling {
                push_labeled_line(&mut out, 4, "Left Sibling", left);
            }
            if let Some(right) = entry.right_sibling {
                push_labeled_line(&mut out, 4, "Right Sibling", right);
            }
            if let Some(child) = entry.child {
                push_labeled_line(&mut out, 4, "Child", child);
            }
            if let Some(created) = entry.created_filetime {
                push_labeled_line(&mut out, 4, "Created", created);
            }
            if let Some(modified) = entry.modified_filetime {
                push_labeled_line(&mut out, 4, "Modified", modified);
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::{format_inspection_text, run};
    use docir_app::{
        test_support::{build_test_cfb, build_test_cfb_with_times},
        BucketCount, DirectoryAnomalySeverity, DirectoryEntry, DirectoryInspection, ParserConfig,
    };
    use std::fs;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn temp_file(name: &str, ext: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_inspect_directory_{name}_{nanos}.{ext}"))
    }

    #[test]
    fn inspect_directory_run_writes_json() {
        let input = temp_file("legacy", "doc");
        let output = temp_file("legacy", "json");
        fs::write(
            &input,
            build_test_cfb(&[("WordDocument", b"doc"), ("VBA/PROJECT", b"meta")]),
        )
        .expect("fixture");

        run(
            input.clone(),
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inspect-directory json");

        let text = fs::read_to_string(&output).expect("output");
        assert!(text.contains("\"container\": \"cfb-ole\""));
        assert!(text.contains("\"role_counts\": ["));
        assert!(text.contains("\"path\": \"WordDocument\""));
        assert!(text.contains("\"entry_index\": 1"));
        assert!(text.contains("\"name_len_raw\": 26"));
        assert!(text.contains("\"object_type_raw\": 2"));
        assert!(text.contains("\"color_flag_raw\": 0"));
        assert!(text.contains("\"left_sibling_raw\": 4294967295"));
        assert!(text.contains("\"child_raw\": 4294967295"));
        assert!(text.contains("\"classification\": \"word-main-stream\""));
        assert!(text.contains("\"anomaly_counts\": ["));
        assert!(text.contains("\"child\":"));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn inspect_directory_run_writes_text() {
        let input = temp_file("legacy_text", "doc");
        let output = temp_file("legacy_text", "txt");
        fs::write(
            &input,
            build_test_cfb_with_times(&[("WordDocument", b"doc")], &[("WordDocument", 11, 22)]),
        )
        .expect("fixture");

        run(
            input.clone(),
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inspect-directory text");

        let text = fs::read_to_string(&output).expect("output");
        assert!(text.contains("Role Counts:"));
        assert!(text.contains("classification:word-main-stream: 1"));
        assert!(text.contains("Anomalies:"));
        assert!(text.contains("free-slot:"));
        assert!(text.contains("Reference Summary:"));
        assert!(text.contains("incoming:"));
        assert!(text.contains("Pointer Summary:"));
        assert!(text.contains("Tree Density Summary:"));
        assert!(text.contains("Reachability Summary:"));
        assert!(text.contains("Incoming Source Summary:"));
        assert!(text.contains("Directory:"));
        assert!(text.contains("stream: WordDocument"));
        assert!(text.contains("Entry Index: 1"));
        assert!(text.contains("Name Length Raw: 26"));
        assert!(text.contains("Object Type Raw: 2"));
        assert!(text.contains("Color Flag Raw: 0"));
        assert!(text.contains("Left Sibling Raw: 4294967295"));
        assert!(text.contains("Child Raw: 4294967295"));
        assert!(text.contains("Classification: word-main-stream"));
        assert!(text.contains("Anomaly Tags:"));
        assert!(
            text.contains("zero-name-length")
                || text.contains("orphaned-entry")
                || text.contains("free-slot")
        );
        assert!(text.contains("Child:"));
        assert!(text.contains("Created: 11"));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn format_inspection_text_renders_expected_fields() {
        let inspection = DirectoryInspection {
            container: "cfb-ole".to_string(),
            entry_count: 1,
            directory_score: "medium".to_string(),
            role_counts: vec![
                BucketCount {
                    bucket: "state:normal".to_string(),
                    count: 1,
                },
                BucketCount {
                    bucket: "classification:word-main-stream".to_string(),
                    count: 1,
                },
            ],
            anomaly_counts: vec![BucketCount {
                bucket: "orphaned-entry".to_string(),
                count: 1,
            }],
            anomaly_catalog: vec![DirectoryAnomalySeverity {
                anomaly: "orphaned-entry".to_string(),
                severity: "medium".to_string(),
            }],
            anomaly_severity_counts: vec![BucketCount {
                bucket: "medium".to_string(),
                count: 1,
            }],
            reference_counts: vec![
                BucketCount {
                    bucket: "incoming:many".to_string(),
                    count: 1,
                },
                BucketCount {
                    bucket: "live-incoming:many".to_string(),
                    count: 1,
                },
            ],
            pointer_counts: vec![
                BucketCount {
                    bucket: "right:present".to_string(),
                    count: 1,
                },
                BucketCount {
                    bucket: "right:dangling".to_string(),
                    count: 1,
                },
            ],
            tree_density_counts: vec![
                BucketCount {
                    bucket: "right:state:normal".to_string(),
                    count: 1,
                },
                BucketCount {
                    bucket: "right:entry-type:stream".to_string(),
                    count: 1,
                },
            ],
            dangling_state_counts: vec![BucketCount {
                bucket: "right:state:normal".to_string(),
                count: 1,
            }],
            self_reference_counts: vec![BucketCount {
                bucket: "self-right-sibling".to_string(),
                count: 1,
            }],
            short_cycle_counts: vec![BucketCount {
                bucket: "sibling-2-cycle".to_string(),
                count: 1,
            }],
            reachability_counts: vec![BucketCount {
                bucket: "live-reachable".to_string(),
                count: 1,
            }],
            incoming_source_counts: vec![BucketCount {
                bucket: "incoming:state:normal".to_string(),
                count: 2,
            }],
            incoming_source_type_counts: vec![
                BucketCount {
                    bucket: "incoming:source-type:root-storage".to_string(),
                    count: 1,
                },
                BucketCount {
                    bucket: "incoming:source-type:storage".to_string(),
                    count: 1,
                },
            ],
            dead_reference_counts: vec![BucketCount {
                bucket: "dead-reference:state:orphaned".to_string(),
                count: 1,
            }],
            fanout_counts: vec![
                BucketCount {
                    bucket: "fanout:0".to_string(),
                    count: 1,
                },
                BucketCount {
                    bucket: "fanout:2".to_string(),
                    count: 1,
                },
            ],
            entries: vec![DirectoryEntry {
                entry_index: 1,
                path: "WordDocument".to_string(),
                entry_type: "stream".to_string(),
                name_len_raw: 26,
                object_type_raw: 2,
                color_flag_raw: 0,
                state: "normal".to_string(),
                classification: "word-main-stream".to_string(),
                anomaly_severity: "medium".to_string(),
                anomaly_tags: vec!["orphaned-entry".to_string()],
                short_cycles: vec!["sibling-2-cycle".to_string()],
                reachable_from_root: true,
                fanout_count: 2,
                incoming_reference_count: 2,
                incoming_normal_reference_count: 2,
                incoming_anomalous_reference_count: 0,
                incoming_from_root_storage_count: 1,
                incoming_from_storage_count: 1,
                incoming_from_stream_count: 0,
                incoming_from: vec!["left:Root Entry#0".to_string(), "right:VBA#2".to_string()],
                size_bytes: 3,
                start_sector: 0,
                left_sibling_raw: u32::MAX,
                right_sibling_raw: 2,
                child_raw: u32::MAX,
                left_sibling: None,
                right_sibling: Some(2),
                child: None,
                created_filetime: None,
                modified_filetime: None,
            }],
        };
        let text = format_inspection_text(&inspection);
        assert!(text.contains("Entries: 1"));
        assert!(text.contains("Directory Score: medium"));
        assert!(text.contains("Role Counts:"));
        assert!(text.contains("classification:word-main-stream: 1"));
        assert!(text.contains("Anomalies:"));
        assert!(text.contains("orphaned-entry: 1"));
        assert!(text.contains("Anomaly Severity Catalog:"));
        assert!(text.contains("orphaned-entry: medium"));
        assert!(text.contains("Anomaly Severity Summary:"));
        assert!(text.contains("medium: 1"));
        assert!(text.contains("Reference Summary:"));
        assert!(text.contains("incoming:many: 1"));
        assert!(text.contains("Pointer Summary:"));
        assert!(text.contains("right:present: 1"));
        assert!(text.contains("Tree Density Summary:"));
        assert!(text.contains("right:state:normal: 1"));
        assert!(text.contains("Dangling By State:"));
        assert!(text.contains("Self References:"));
        assert!(text.contains("Short Cycles:"));
        assert!(text.contains("Reachability Summary:"));
        assert!(text.contains("Incoming Source Summary:"));
        assert!(text.contains("Incoming Source Types:"));
        assert!(text.contains("Dead But Referenced:"));
        assert!(text.contains("Fanout Summary:"));
        assert!(text.contains("Entry Index: 1"));
        assert!(text.contains("Name Length Raw: 26"));
        assert!(text.contains("Object Type Raw: 2"));
        assert!(text.contains("Color Flag Raw: 0"));
        assert!(text.contains("State: normal"));
        assert!(text.contains("Anomaly Severity: medium"));
        assert!(text.contains("Short Cycles: sibling-2-cycle"));
        assert!(text.contains("Reachable From Root: true"));
        assert!(text.contains("Fanout: 2"));
        assert!(text.contains("Incoming References: 2"));
        assert!(text.contains("Incoming From Normal: 2"));
        assert!(text.contains("Incoming From Anomalous: 0"));
        assert!(text.contains("Incoming From Root Storage: 1"));
        assert!(text.contains("Incoming From Storage: 1"));
        assert!(text.contains("Incoming From Stream: 0"));
        assert!(text.contains("Incoming From: left:Root Entry#0, right:VBA#2"));
        assert!(text.contains("Sector: 0"));
        assert!(text.contains("Anomaly Tags: orphaned-entry"));
        assert!(text.contains("Left Sibling Raw: 4294967295"));
        assert!(text.contains("Right Sibling Raw: 2"));
        assert!(text.contains("Right Sibling: 2"));
    }
}
