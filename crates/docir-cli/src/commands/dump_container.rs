//! Dump low-level container entries for analyst inspection.

use anyhow::Result;
use docir_app::{ContainerDump, ParserConfig};
use std::path::PathBuf;

use crate::commands::util::{build_app, push_bullet_line, push_labeled_line, run_dual_output};

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let app = build_app(parser_config);
    let (parsed, bytes) = app.parse_file_with_bytes(&input)?;
    let dump = app.build_container_dump(&parsed, &bytes)?;
    run_dual_output(
        &dump,
        "container",
        json,
        pretty,
        output,
        format_container_text,
    )
}

fn format_container_text(dump: &ContainerDump) -> String {
    let mut out = String::new();
    push_labeled_line(&mut out, 0, "Format", &dump.document_format);
    push_labeled_line(
        &mut out,
        0,
        "Container",
        format!("{:?}", dump.container_kind),
    );
    push_labeled_line(&mut out, 0, "Entries", dump.entry_count);

    if !dump.entries.is_empty() {
        out.push_str("\nEntries:\n");
        for entry in &dump.entries {
            push_bullet_line(&mut out, 2, container_entry_label(entry.kind), &entry.path);
            if let Some(size_bytes) = entry.size_bytes {
                push_labeled_line(&mut out, 4, "Size", format!("{size_bytes} bytes"));
            }
            if let Some(start_sector) = entry.start_sector {
                push_labeled_line(&mut out, 4, "CFB Sector", start_sector);
            }
            if let Some(details) = entry.details.as_deref() {
                push_labeled_line(&mut out, 4, "Classification", details);
            }
            if let Some(created) = entry.created_filetime {
                push_labeled_line(&mut out, 4, "CFB Created", created);
            }
            if let Some(modified) = entry.modified_filetime {
                push_labeled_line(&mut out, 4, "CFB Modified", modified);
            }
        }
    }

    out
}

fn container_entry_label(kind: docir_app::ContainerEntryKind) -> &'static str {
    match kind {
        docir_app::ContainerEntryKind::ZipEntry => "ZIP Entry",
        docir_app::ContainerEntryKind::CfbRootStorage => "CFB Root Storage",
        docir_app::ContainerEntryKind::CfbStorage => "CFB Storage",
        docir_app::ContainerEntryKind::CfbStream => "CFB Stream",
        docir_app::ContainerEntryKind::RtfDocument => "RTF Document",
        docir_app::ContainerEntryKind::RtfEmbeddedObject => "RTF Embedded Object",
    }
}

#[cfg(test)]
mod tests {
    use super::{format_container_text, run};
    use crate::test_support;
    use docir_app::{
        test_support::build_test_cfb, ContainerDump, ContainerEntry, ContainerEntryKind,
        ContainerKind, ParserConfig,
    };
    use std::fs;

    #[test]
    fn dump_container_run_writes_json() {
        let output = test_support::temp_file("json", "json");
        run(
            test_support::fixture("minimal.docx"),
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("dump container json");
        let text = fs::read_to_string(&output).expect("dump container output");
        assert!(text.contains("\"container\""));
        let _ = fs::remove_file(output);
    }

    #[test]
    fn dump_container_run_writes_text() {
        let output = test_support::temp_file("text", "json");
        run(
            test_support::fixture("minimal.docx"),
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("dump container text");
        let text = fs::read_to_string(&output).expect("dump container output");
        assert!(text.contains("Entries:"));
        assert!(text.contains("word/document.xml"));
        let _ = fs::remove_file(output);
    }

    #[test]
    fn dump_container_run_writes_text_for_legacy_cfb_fixture() {
        let input = test_support::temp_file("legacy_text", "doc");
        fs::write(
            &input,
            build_test_cfb(&[
                ("WordDocument", b"doc"),
                ("VBA/PROJECT", b"ID=\"VBAProject\""),
                ("ObjectPool/1/Ole10Native", b"payload"),
            ]),
        )
        .expect("write legacy fixture");
        let output = test_support::temp_file("legacy_text_out", "json");
        run(
            input.clone(),
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("dump legacy container text");
        let text = fs::read_to_string(&output).expect("dump container output");
        assert!(text.contains("CFB Stream: VBA/PROJECT"));
        assert!(text.contains("Classification: vba-project-metadata"));
        assert!(text.contains("CFB Stream: ObjectPool/1/Ole10Native"));
        assert!(text.contains("Classification: ole-native-payload"));
        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn format_container_text_highlights_vba_and_embedded_object_entries() {
        let dump = ContainerDump {
            document_format: "doc".to_string(),
            container_kind: ContainerKind::CfbOle,
            entry_count: 5,
            entries: vec![
                ContainerEntry {
                    kind: ContainerEntryKind::CfbRootStorage,
                    path: "Root Entry".to_string(),
                    size_bytes: Some(0),
                    start_sector: None,
                    created_filetime: None,
                    modified_filetime: None,
                    details: Some("root-storage".to_string()),
                },
                ContainerEntry {
                    kind: ContainerEntryKind::CfbStorage,
                    path: "VBA".to_string(),
                    size_bytes: Some(0),
                    start_sector: Some(1),
                    created_filetime: Some(100),
                    modified_filetime: Some(200),
                    details: None,
                },
                ContainerEntry {
                    kind: ContainerEntryKind::CfbStream,
                    path: "VBA/PROJECT".to_string(),
                    size_bytes: Some(64),
                    start_sector: Some(2),
                    created_filetime: None,
                    modified_filetime: None,
                    details: Some("vba-project-metadata".to_string()),
                },
                ContainerEntry {
                    kind: ContainerEntryKind::CfbStorage,
                    path: "ObjectPool/1".to_string(),
                    size_bytes: Some(0),
                    start_sector: Some(3),
                    created_filetime: None,
                    modified_filetime: None,
                    details: Some("embedded-object-storage".to_string()),
                },
                ContainerEntry {
                    kind: ContainerEntryKind::CfbStream,
                    path: "ObjectPool/1/Ole10Native".to_string(),
                    size_bytes: Some(128),
                    start_sector: Some(4),
                    created_filetime: Some(123),
                    modified_filetime: Some(456),
                    details: Some("ole-native-payload".to_string()),
                },
            ],
        };

        let text = format_container_text(&dump);
        assert!(text.contains("Format: doc"));
        assert!(text.contains("Container: CfbOle"));
        assert!(text.contains("Entries: 5"));
        assert!(text.contains("- CFB Root Storage: Root Entry"));
        assert!(text.contains("Classification: root-storage"));
        assert!(text.contains("- CFB Storage: VBA"));
        assert!(text.contains("CFB Sector: 1"));
        assert!(text.contains("CFB Created: 100"));
        assert!(text.contains("CFB Modified: 200"));
        assert!(text.contains("- CFB Stream: VBA/PROJECT"));
        assert!(text.contains("Classification: vba-project-metadata"));
        assert!(text.contains("- CFB Storage: ObjectPool/1"));
        assert!(text.contains("Classification: embedded-object-storage"));
        assert!(text.contains("- CFB Stream: ObjectPool/1/Ole10Native"));
        assert!(text.contains("Classification: ole-native-payload"));
        assert!(text.contains("Size: 128 bytes"));
    }
}
