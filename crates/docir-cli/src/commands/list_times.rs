//! List CFB storage and stream FILETIMEs.

use anyhow::Result;
use docir_app::{list_times_path, ParserConfig, TimeListing};
use serde::Serialize;
use std::path::PathBuf;

use crate::commands::util::{
    push_bullet_line, push_labeled_line, write_json_output, write_text_output,
};

#[derive(Debug, Serialize)]
struct ListTimesResult {
    listing: TimeListing,
}

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let listing = list_times_path(&input, parser_config)?;

    if json {
        return write_json_output(&ListTimesResult { listing }, pretty, output);
    }

    let text = format_listing_text(&listing);
    write_text_output(&text, output)
}

fn format_listing_text(listing: &TimeListing) -> String {
    let mut out = String::new();
    push_labeled_line(&mut out, 0, "Container", &listing.container);
    push_labeled_line(&mut out, 0, "Entries", listing.entry_count);
    if !listing.entries.is_empty() {
        out.push_str("\nTimestamps:\n");
        for entry in &listing.entries {
            push_bullet_line(&mut out, 2, &entry.entry_type, &entry.path);
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
    use super::{format_listing_text, run};
    use docir_app::{
        test_support::build_test_cfb_with_times, ParserConfig, TimeEntry, TimeListing,
    };
    use std::fs;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn temp_file(name: &str, ext: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_list_times_{name}_{nanos}.{ext}"))
    }

    #[test]
    fn list_times_run_writes_json() {
        let input = temp_file("legacy", "doc");
        let output = temp_file("legacy", "json");
        fs::write(
            &input,
            build_test_cfb_with_times(&[("WordDocument", b"doc")], &[("WordDocument", 10, 20)]),
        )
        .expect("fixture");

        run(
            input.clone(),
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("list-times json");

        let text = fs::read_to_string(&output).expect("output");
        assert!(text.contains("\"container\": \"cfb-ole\""));
        assert!(text.contains("\"path\": \"WordDocument\""));
        assert!(text.contains("\"created_filetime\": 10"));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn list_times_run_writes_text() {
        let input = temp_file("legacy_text", "doc");
        let output = temp_file("legacy_text", "txt");
        fs::write(
            &input,
            build_test_cfb_with_times(
                &[("WordDocument", b"doc"), ("VBA/PROJECT", b"meta")],
                &[("VBA/PROJECT", 30, 40)],
            ),
        )
        .expect("fixture");

        run(
            input.clone(),
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("list-times text");

        let text = fs::read_to_string(&output).expect("output");
        assert!(text.contains("Container: cfb-ole"));
        assert!(text.contains("Timestamps:"));
        assert!(text.contains("stream: VBA/PROJECT"));
        assert!(text.contains("Created: 30"));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn format_listing_text_renders_expected_fields() {
        let listing = TimeListing {
            container: "cfb-ole".to_string(),
            entry_count: 1,
            entries: vec![TimeEntry {
                path: "WordDocument".to_string(),
                entry_type: "stream".to_string(),
                created_filetime: Some(1),
                modified_filetime: Some(2),
            }],
        };
        let text = format_listing_text(&listing);
        assert!(text.contains("Entries: 1"));
        assert!(text.contains("stream: WordDocument"));
        assert!(text.contains("Modified: 2"));
    }
}
