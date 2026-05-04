//! Probe the real format/container of an input file.

use anyhow::Result;
use docir_app::{probe_format_path, FormatProbe, ParserConfig};
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
    let probe = probe_format_path(&input, parser_config)?;
    run_dual_output(&probe, "probe", json, pretty, output, format_probe_text)
}

fn format_probe_text(probe: &FormatProbe) -> String {
    let mut out = String::new();
    push_labeled_line(&mut out, 0, "Format", &probe.format);
    push_labeled_line(&mut out, 0, "Container", &probe.container);
    push_labeled_line(&mut out, 0, "Family", &probe.family);
    push_labeled_line(
        &mut out,
        0,
        "Suggested Extension",
        &probe.suggested_extension,
    );
    push_labeled_line(&mut out, 0, "Confidence", &probe.confidence);
    if !probe.signals.is_empty() {
        out.push_str("\nSignals:\n");
        for signal in &probe.signals {
            push_bullet_line(&mut out, 2, "Signal", signal);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::run;
    use crate::cli::JsonOutputOpts;
    use crate::test_support;
    use docir_app::{test_support::build_test_cfb, FormatProbe, ParserConfig};
    use std::fs;

    #[test]
    fn probe_format_run_writes_json_for_pdf() {
        let input = test_support::temp_file("pdf", "bin");
        let output = test_support::temp_file("pdf_out", "json");
        fs::write(&input, b"%PDF-1.7\n").expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: true,
                pretty: true,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("probe json");

        let text = fs::read_to_string(&output).expect("output");
        assert!(text.contains("\"format\": \"pdf\""));
        assert!(text.contains("\"container\": \"raw-binary\""));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn probe_format_run_writes_text_for_docx() {
        let input = test_support::temp_file("docx", "docx");
        let output = test_support::temp_file("docx_out", "txt");
        test_support::write_docx(&input);

        run(
            input.clone(),
            JsonOutputOpts {
                json: false,
                pretty: false,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("probe text");

        let text = fs::read_to_string(&output).expect("output");
        assert!(text.contains("Format: docx"));
        assert!(text.contains("Container: zip-ooxml"));
        assert!(text.contains("Signals:"));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn probe_format_run_writes_text_for_legacy_doc() {
        let input = test_support::temp_file("legacy_doc", "doc");
        let output = test_support::temp_file("legacy_doc_out", "txt");
        fs::write(&input, build_test_cfb(&[("WordDocument", b"doc")])).expect("fixture");

        run(
            input.clone(),
            JsonOutputOpts {
                json: false,
                pretty: false,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("probe legacy doc");

        let text = fs::read_to_string(&output).expect("output");
        assert!(text.contains("Format: doc"));
        assert!(text.contains("Container: cfb-ole"));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn format_probe_text_renders_expected_fields() {
        let probe = FormatProbe {
            format: "docx".to_string(),
            container: "zip-ooxml".to_string(),
            family: "word-processing".to_string(),
            suggested_extension: "docx".to_string(),
            confidence: "high".to_string(),
            signals: vec!["zip-signature".to_string()],
        };
        let text = super::format_probe_text(&probe);
        assert!(text.contains("Family: word-processing"));
        assert!(text.contains("Suggested Extension: docx"));
        assert!(text.contains("Signal: zip-signature"));
    }
}
