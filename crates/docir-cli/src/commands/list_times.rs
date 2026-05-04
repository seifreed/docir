//! List CFB storage and stream FILETIMEs.

use anyhow::Result;
use docir_app::{list_times_path, ParserConfig, TimeListing};
use std::path::PathBuf;

use crate::cli::JsonOutputOpts;
use crate::commands::util::{push_bullet_line, push_labeled_line, run_dual_output};

/// Public API entrypoint: run.
pub fn run(input: PathBuf, opts: JsonOutputOpts, parser_config: &ParserConfig) -> Result<()> {
    let JsonOutputOpts {
        json,
        pretty,
        output,
    } = opts;
    let listing = list_times_path(&input, parser_config)?;
    run_dual_output(
        &listing,
        "listing",
        json,
        pretty,
        output,
        format_listing_text,
    )
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
    use crate::cli::JsonOutputOpts;
    use crate::test_support;
    use docir_app::{
        test_support::build_test_cfb_with_times, ParserConfig, TimeEntry, TimeListing,
    };
    use std::fs;

    #[test]
    fn list_times_run_writes_json() {
        let input = test_support::temp_file("legacy", "doc");
        let output = test_support::temp_file("legacy", "json");
        fs::write(
            &input,
            build_test_cfb_with_times(&[("WordDocument", b"doc")], &[("WordDocument", 10, 20)]),
        )
        .expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: true,
                pretty: true,
                output: Some(output.clone()),
            },
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
        let input = test_support::temp_file("legacy_text", "doc");
        let output = test_support::temp_file("legacy_text", "txt");
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
            JsonOutputOpts {
                json: false,
                pretty: false,
                output: Some(output.clone()),
            },
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
