//! Extract embedded SWF/Flash payloads from containers.

use anyhow::{Context, Result};
use docir_app::{extract_flash_path, FlashExtractionReport, ParserConfig};
use serde::Serialize;
use std::fs;
use std::path::{Path, PathBuf};

use crate::commands::util::{
    push_bullet_line, push_labeled_line, write_json_output, write_text_output,
};

#[derive(Debug, Serialize)]
struct ExtractFlashResult {
    report: FlashExtractionReport,
}

pub fn run(
    input: PathBuf,
    out: Option<PathBuf>,
    overwrite: bool,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let mut report = extract_flash_path(&input, parser_config)?;

    if let Some(out_dir) = out.as_ref() {
        prepare_output_dir(out_dir, overwrite)?;
        for (index, object) in report.objects.iter_mut().enumerate() {
            let file_name = format!("flash_{}.swf", index + 1);
            let path = out_dir.join(&file_name);
            fs::write(&path, &object.data)
                .with_context(|| format!("Failed to write {}", path.display()))?;
            object.output_path = Some(format!("payloads/{}", file_name));
        }
    }

    if json {
        return write_json_output(&ExtractFlashResult { report }, pretty, output);
    }

    let text = format_report_text(&report);
    write_text_output(&text, output)
}

fn prepare_output_dir(path: &Path, overwrite: bool) -> Result<()> {
    if path.exists() {
        if !overwrite {
            anyhow::bail!("Output directory {} already exists", path.display());
        }
    } else {
        fs::create_dir_all(path).with_context(|| format!("Failed to create {}", path.display()))?;
    }
    Ok(())
}

fn format_report_text(report: &FlashExtractionReport) -> String {
    let mut out = String::new();
    push_labeled_line(&mut out, 0, "Container", &report.container);
    push_labeled_line(&mut out, 0, "Objects", report.object_count);
    if !report.objects.is_empty() {
        out.push_str("\nFlash Objects:\n");
        for object in &report.objects {
            push_bullet_line(
                &mut out,
                2,
                &object.source_path,
                format!("{} @{}", object.signature, object.offset),
            );
            push_labeled_line(&mut out, 4, "Compression", &object.compression);
            push_labeled_line(&mut out, 4, "Version", object.version);
            push_labeled_line(&mut out, 4, "Declared Size", object.declared_size);
            push_labeled_line(&mut out, 4, "Extracted Size", object.extracted_size);
            push_labeled_line(&mut out, 4, "Truncated", object.truncated);
            push_labeled_line(&mut out, 4, "SHA-256", &object.sha256);
            if let Some(path) = &object.output_path {
                push_labeled_line(&mut out, 4, "Output", path);
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::{format_report_text, run};
    use docir_app::{FlashExtractionReport, FlashObject, ParserConfig};
    use std::fs;
    use std::io::Write;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn temp_file(name: &str, ext: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_extract_flash_{name}_{nanos}.{ext}"))
    }

    fn temp_dir(name: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_extract_flash_{name}_{nanos}"))
    }

    fn swf(signature: &[u8; 3], version: u8, body: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(signature);
        out.push(version);
        out.extend_from_slice(&((body.len() + 8) as u32).to_le_bytes());
        out.extend_from_slice(body);
        out
    }

    #[test]
    fn extract_flash_run_writes_json_for_zip_payload() {
        let input = temp_file("flash_zip", "docx");
        let output = temp_file("flash_zip", "json");
        let mut cursor = std::io::Cursor::new(Vec::<u8>::new());
        {
            let mut zip = zip::ZipWriter::new(&mut cursor);
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("word/media/movie.bin", options)
                .expect("start");
            zip.write_all(&swf(b"FWS", 8, b"payload")).expect("write");
            zip.finish().expect("finish");
        }
        fs::write(&input, cursor.into_inner()).expect("fixture");

        run(
            input,
            None,
            false,
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("extract-flash json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"container\": \"zip\""));
        assert!(text.contains("\"signature\": \"FWS\""));
        assert!(text.contains("\"source_path\": \"word/media/movie.bin\""));
    }

    #[test]
    fn extract_flash_run_writes_payload_to_out_dir() {
        let input = temp_file("flash_out", "docx");
        let out_dir = temp_dir("flash_out");
        let mut cursor = std::io::Cursor::new(Vec::<u8>::new());
        {
            let mut zip = zip::ZipWriter::new(&mut cursor);
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("word/media/movie.bin", options)
                .expect("start");
            zip.write_all(&swf(b"CWS", 10, b"payload")).expect("write");
            zip.finish().expect("finish");
        }
        fs::write(&input, cursor.into_inner()).expect("fixture");

        run(
            input,
            Some(out_dir.clone()),
            false,
            false,
            false,
            None,
            &ParserConfig::default(),
        )
        .expect("extract-flash out");

        let payload = out_dir.join("flash_1.swf");
        let bytes = fs::read(&payload).expect("payload");
        assert!(bytes.starts_with(b"CWS"));
    }

    #[test]
    fn extract_flash_run_writes_json_for_cfb_payload() {
        let input = temp_file("flash_cfb", "doc");
        let output = temp_file("flash_cfb", "json");
        fs::write(
            &input,
            docir_app::test_support::build_test_cfb(&[(
                "ObjectPool/1/Ole10Native",
                &swf(b"ZWS", 13, b"payload"),
            )]),
        )
        .expect("fixture");

        run(
            input,
            None,
            false,
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("extract-flash cfb json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"container\": \"cfb\""));
        assert!(text.contains("\"signature\": \"ZWS\""));
        assert!(text.contains("\"source_path\": \"ObjectPool/1/Ole10Native\""));
    }

    #[test]
    fn extract_flash_run_reports_truncated_cws_from_zip_entry() {
        let input = temp_file("flash_truncated_zip", "pptx");
        let output = temp_file("flash_truncated_zip", "json");
        let mut payload = swf(b"CWS", 10, b"payload");
        payload[4..8].copy_from_slice(&128u32.to_le_bytes());
        let mut cursor = std::io::Cursor::new(Vec::<u8>::new());
        {
            let mut zip = zip::ZipWriter::new(&mut cursor);
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("ppt/media/flash.bin", options)
                .expect("start");
            zip.write_all(&payload).expect("write");
            zip.finish().expect("finish");
        }
        fs::write(&input, cursor.into_inner()).expect("fixture");

        run(
            input,
            None,
            false,
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("extract-flash truncated zip json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"container\": \"zip\""));
        assert!(text.contains("\"signature\": \"CWS\""));
        assert!(text.contains("\"truncated\": true"));
        assert!(text.contains("\"source_path\": \"ppt/media/flash.bin\""));
    }

    #[test]
    fn extract_flash_run_writes_text_for_truncated_cws_zip_entry() {
        let input = temp_file("flash_truncated_zip_text", "pptx");
        let output = temp_file("flash_truncated_zip_text", "txt");
        let mut payload = swf(b"CWS", 11, b"payload");
        payload[4..8].copy_from_slice(&128u32.to_le_bytes());
        let mut cursor = std::io::Cursor::new(Vec::<u8>::new());
        {
            let mut zip = zip::ZipWriter::new(&mut cursor);
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("ppt/media/flash.bin", options)
                .expect("start");
            zip.write_all(&payload).expect("write");
            zip.finish().expect("finish");
        }
        fs::write(&input, cursor.into_inner()).expect("fixture");

        run(
            input,
            None,
            false,
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("extract-flash truncated zip text");

        let text = fs::read_to_string(&output).expect("text output");
        assert!(text.contains("Container: zip"));
        assert!(text.contains("ppt/media/flash.bin"));
        assert!(text.contains("Compression: zlib"));
        assert!(text.contains("Truncated: true"));
    }

    #[test]
    fn extract_flash_run_reports_raw_zws_payload_in_text() {
        let input = temp_file("flash_raw_zws", "bin");
        let output = temp_file("flash_raw_zws", "txt");
        fs::write(&input, swf(b"ZWS", 13, b"payload")).expect("fixture");

        run(
            input,
            None,
            false,
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("extract-flash raw text");

        let text = fs::read_to_string(&output).expect("text output");
        assert!(text.contains("Container: raw-binary"));
        assert!(text.contains("input"));
        assert!(text.contains("Compression: lzma"));
        assert!(text.contains("ZWS @0"));
    }

    #[test]
    fn extract_flash_run_reports_truncated_cws_from_cfb_text() {
        let input = temp_file("flash_cfb_truncated", "doc");
        let output = temp_file("flash_cfb_truncated", "txt");
        let mut payload = swf(b"CWS", 9, b"payload");
        payload[4..8].copy_from_slice(&256u32.to_le_bytes());
        fs::write(
            &input,
            docir_app::test_support::build_test_cfb(&[("ObjectPool/2/Package", &payload)]),
        )
        .expect("fixture");

        run(
            input,
            None,
            false,
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("extract-flash cfb truncated text");

        let text = fs::read_to_string(&output).expect("text output");
        assert!(text.contains("Container: cfb"));
        assert!(text.contains("ObjectPool/2/Package"));
        assert!(text.contains("Compression: zlib"));
        assert!(text.contains("Truncated: true"));
    }

    #[test]
    fn extract_flash_run_reports_fws_from_zip_json() {
        let input = temp_file("flash_fws_zip", "docx");
        let output = temp_file("flash_fws_zip", "json");
        let mut cursor = std::io::Cursor::new(Vec::<u8>::new());
        {
            let mut zip = zip::ZipWriter::new(&mut cursor);
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("word/embeddings/flash.bin", options)
                .expect("start");
            zip.write_all(&swf(b"FWS", 8, b"payload")).expect("write");
            zip.finish().expect("finish");
        }
        fs::write(&input, cursor.into_inner()).expect("fixture");

        run(
            input,
            None,
            false,
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("extract-flash fws zip json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"container\": \"zip\""));
        assert!(text.contains("\"signature\": \"FWS\""));
        assert!(text.contains("\"source_path\": \"word/embeddings/flash.bin\""));
    }

    #[test]
    fn extract_flash_run_reports_zws_from_cfb_json() {
        let input = temp_file("flash_zws_cfb", "doc");
        let output = temp_file("flash_zws_cfb", "json");
        fs::write(
            &input,
            docir_app::test_support::build_test_cfb(&[(
                "ObjectPool/3/Ole10Native",
                &swf(b"ZWS", 13, b"payload"),
            )]),
        )
        .expect("fixture");

        run(
            input,
            None,
            false,
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("extract-flash zws cfb json");

        let text = fs::read_to_string(&output).expect("json output");
        assert!(text.contains("\"container\": \"cfb\""));
        assert!(text.contains("\"signature\": \"ZWS\""));
        assert!(text.contains("\"compression\": \"lzma\""));
    }

    #[test]
    fn extract_flash_run_reports_cws_from_zip_text() {
        let input = temp_file("flash_cws_zip_text", "docx");
        let output = temp_file("flash_cws_zip_text", "txt");
        let mut cursor = std::io::Cursor::new(Vec::<u8>::new());
        {
            let mut zip = zip::ZipWriter::new(&mut cursor);
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("word/media/flash.bin", options)
                .expect("start");
            zip.write_all(&swf(b"CWS", 9, b"payload")).expect("write");
            zip.finish().expect("finish");
        }
        fs::write(&input, cursor.into_inner()).expect("fixture");

        run(
            input,
            None,
            false,
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("extract-flash cws zip text");

        let text = fs::read_to_string(&output).expect("text output");
        assert!(text.contains("Container: zip"));
        assert!(text.contains("word/media/flash.bin"));
        assert!(text.contains("Compression: zlib"));
        assert!(text.contains("CWS @0"));
    }

    #[test]
    fn format_report_text_shows_truncated_payloads() {
        let report = FlashExtractionReport {
            container: "raw-binary".to_string(),
            object_count: 1,
            objects: vec![FlashObject {
                source_path: "input".to_string(),
                offset: 4,
                signature: "FWS".to_string(),
                compression: "none".to_string(),
                version: 8,
                declared_size: 999,
                extracted_size: 20,
                truncated: true,
                sha256: "cafebabe".to_string(),
                data: b"truncated-flash".to_vec(),
                output_path: None,
            }],
        };
        let text = format_report_text(&report);
        assert!(text.contains("Truncated: true"));
        assert!(text.contains("Declared Size: 999"));
    }

    #[test]
    fn format_report_text_renders_expected_fields() {
        let report = FlashExtractionReport {
            container: "zip".to_string(),
            object_count: 1,
            objects: vec![FlashObject {
                source_path: "word/media/movie.bin".to_string(),
                offset: 10,
                signature: "FWS".to_string(),
                compression: "none".to_string(),
                version: 9,
                declared_size: 24,
                extracted_size: 24,
                truncated: false,
                sha256: "deadbeef".to_string(),
                data: b"flash-payload".to_vec(),
                output_path: Some("payloads/flash_1.swf".to_string()),
            }],
        };
        let text = format_report_text(&report);
        assert!(text.contains("Container: zip"));
        assert!(text.contains("word/media/movie.bin"));
        assert!(text.contains("Compression: none"));
        assert!(text.contains("Output: payloads/flash_1.swf"));
    }
}
