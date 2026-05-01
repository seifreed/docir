//! Probe the real format/container of an input file.

use anyhow::Result;
use docir_app::{probe_format_path, FormatProbe, ParserConfig};
use serde::Serialize;
use std::path::PathBuf;

use crate::commands::util::{
    push_bullet_line, push_labeled_line, write_json_output, write_text_output,
};

#[derive(Debug, Serialize)]
struct ProbeFormatResult {
    probe: FormatProbe,
}

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let probe = probe_format_path(&input, parser_config)?;

    if json {
        return write_json_output(&ProbeFormatResult { probe }, pretty, output);
    }

    let text = format_probe_text(&probe);
    write_text_output(&text, output)
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
    use docir_app::{test_support::build_test_cfb, FormatProbe, ParserConfig};
    use std::fs;
    use std::io::Write;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};
    use zip::write::FileOptions;

    fn temp_file(name: &str, ext: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_probe_format_{name}_{nanos}.{ext}"))
    }

    fn temp_output(name: &str, ext: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_probe_format_out_{name}_{nanos}.{ext}"))
    }

    fn write_docx(path: &PathBuf) {
        let file = fs::File::create(path).expect("create docx");
        let mut zip = zip::ZipWriter::new(file);
        let options = FileOptions::<()>::default();
        let content_types = r#"
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml"
                ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>"#;
        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.trim().as_bytes()).unwrap();
        zip.add_directory("word/", options).unwrap();
        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(b"<w:document/>").unwrap();
        zip.finish().unwrap();
    }

    #[test]
    fn probe_format_run_writes_json_for_pdf() {
        let input = temp_file("pdf", "bin");
        let output = temp_output("pdf", "json");
        fs::write(&input, b"%PDF-1.7\n").expect("fixture");

        run(
            input.clone(),
            true,
            true,
            Some(output.clone()),
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
        let input = temp_file("docx", "docx");
        let output = temp_output("docx", "txt");
        write_docx(&input);

        run(
            input.clone(),
            false,
            false,
            Some(output.clone()),
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
        let input = temp_file("legacy_doc", "doc");
        let output = temp_output("legacy_doc", "txt");
        fs::write(&input, build_test_cfb(&[("WordDocument", b"doc")])).expect("fixture");

        run(
            input.clone(),
            false,
            false,
            Some(output.clone()),
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
