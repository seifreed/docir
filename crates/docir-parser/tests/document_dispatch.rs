use docir_core::types::DocumentFormat;
use docir_parser::{DocumentParser, ParseError};
use std::fs;
use std::io::{Cursor, Write};
use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};

fn build_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::SimpleFileOptions::default();
    for (path, data) in entries {
        zip.start_file(path, options).expect("zip start_file");
        zip.write_all(data).expect("zip write_all");
    }
    zip.finish().expect("zip finish").into_inner()
}

fn temp_file_path(name: &str, ext: &str) -> PathBuf {
    let mut path = std::env::temp_dir();
    let ts = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("clock before unix epoch")
        .as_nanos();
    path.push(format!("docir_{name}_{ts}.{ext}"));
    path
}

#[test]
fn parse_reader_detects_rtf_and_parse_reader_with_bytes_returns_raw_data() {
    let parser = DocumentParser::new();
    let rtf_bytes = b"{\\rtf1\\ansi Integration Dispatch}".to_vec();

    let parsed = parser
        .parse_reader(Cursor::new(rtf_bytes.clone()))
        .expect("parse rtf via parse_reader");
    assert_eq!(parsed.format, DocumentFormat::Rtf);
    assert!(parsed.document().is_some());

    let (parsed_with_bytes, returned) = parser
        .parse_reader_with_bytes(Cursor::new(rtf_bytes.clone()))
        .expect("parse rtf via parse_reader_with_bytes");
    assert_eq!(parsed_with_bytes.format, DocumentFormat::Rtf);
    assert_eq!(returned, rtf_bytes);
}

#[test]
fn parse_file_with_bytes_detects_rtf_and_returns_original_bytes() {
    let parser = DocumentParser::new();
    let path = temp_file_path("dispatch_rtf", "rtf");
    let rtf_bytes = b"{\\rtf1\\ansi File Dispatch}".to_vec();
    fs::write(&path, &rtf_bytes).expect("write temp rtf");

    let (parsed, returned) = parser
        .parse_file_with_bytes(&path)
        .expect("parse_file_with_bytes rtf");
    assert_eq!(parsed.format, DocumentFormat::Rtf);
    assert_eq!(returned, rtf_bytes);

    fs::remove_file(path).ok();
}

#[test]
fn parse_reader_rejects_unknown_zip_markers_in_normal_and_with_bytes_paths() {
    let parser = DocumentParser::new();
    let unknown_zip = build_zip(&[("data.bin", b"payload")]);

    let err = parser
        .parse_reader(Cursor::new(unknown_zip.clone()))
        .expect_err("zip without ooxml/odf markers must fail");
    match err {
        ParseError::UnsupportedFormat(message) => {
            assert!(message.contains("missing [Content_Types].xml and mimetype"));
        }
        other => panic!("unexpected error: {other:?}"),
    }

    let err_with_bytes = parser
        .parse_reader_with_bytes(Cursor::new(unknown_zip))
        .expect_err("with_bytes should fail for unknown zip");
    match err_with_bytes {
        ParseError::UnsupportedFormat(message) => {
            assert!(message.contains("missing [Content_Types].xml and mimetype"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn parse_file_with_bytes_rejects_unknown_zip_markers() {
    let parser = DocumentParser::new();
    let path = temp_file_path("dispatch_unknown_zip", "zip");
    let unknown_zip = build_zip(&[("payload.bin", b"raw")]);
    fs::write(&path, unknown_zip).expect("write temp zip");

    let err = parser
        .parse_file_with_bytes(&path)
        .expect_err("unknown zip should fail");
    match err {
        ParseError::UnsupportedFormat(message) => {
            assert!(message.contains("missing [Content_Types].xml and mimetype"));
        }
        other => panic!("unexpected error: {other:?}"),
    }

    fs::remove_file(path).ok();
}

#[test]
fn parse_reader_ole_header_dispatches_to_hwp_path() {
    let parser = DocumentParser::new();
    let mut bytes = vec![0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    bytes.extend_from_slice(&[0u8; 512]);

    let err = parser
        .parse_reader(Cursor::new(bytes))
        .expect_err("synthetic ole header should fail in hwp parser path");

    match err {
        ParseError::UnsupportedFormat(message) => {
            assert!(
                !message.contains("Unknown package format"),
                "error indicates format detection failed before HWP dispatch: {message}"
            );
        }
        ParseError::InvalidFormat(_)
        | ParseError::InvalidStructure(_)
        | ParseError::MissingPart(_)
        | ParseError::InvalidZip(_) => {}
        other => panic!("unexpected error type: {other:?}"),
    }
}
