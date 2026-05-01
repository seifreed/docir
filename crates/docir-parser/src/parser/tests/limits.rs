use super::super::*;
use super::helpers::create_minimal_docx;
use std::io::Write;
use zip::write::FileOptions;

#[test]
fn test_document_parser_enforces_max_input_size() {
    let config = ParserConfig {
        max_input_size: 32,
        ..ParserConfig::default()
    };
    let parser = DocumentParser::with_config(config);
    let data = vec![b'A'; 128];
    let err = parser
        .parse_reader(std::io::Cursor::new(data))
        .expect_err("expected size limit error");
    assert!(matches!(err, ParseError::ResourceLimit(_)));
}

#[test]
fn test_document_parser_rejects_non_zip_non_ole_input() {
    let parser = DocumentParser::new();
    let err = parser
        .parse_reader(std::io::Cursor::new(b"plain-text".to_vec()))
        .expect_err("non-zip/non-ole should fail");
    match err {
        ParseError::UnsupportedFormat(msg) => {
            assert!(msg.contains("not OLE/CFB or ZIP"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_document_parser_rejects_zip_without_ooxml_or_odf_markers() {
    let mut path = std::env::temp_dir();
    path.push("docir_unknown_zip_format_test.zip");
    let file = std::fs::File::create(&path).expect("create zip");
    let mut zip = zip::ZipWriter::new(file);
    let options = FileOptions::<()>::default();
    zip.start_file("data.bin", options).expect("zip entry");
    zip.write_all(b"payload").expect("zip data");
    zip.finish().expect("zip finish");

    let parser = DocumentParser::new();
    let err = parser
        .parse_file(&path)
        .expect_err("zip without markers should fail");
    match err {
        ParseError::UnsupportedFormat(msg) => {
            assert!(msg.contains("missing [Content_Types].xml and mimetype"));
        }
        other => panic!("unexpected error: {other:?}"),
    }

    std::fs::remove_file(path).ok();
}

#[test]
fn test_parse_file_and_reader_with_bytes_return_raw_data() {
    let path = create_minimal_docx(true);
    let parser = DocumentParser::new();

    let (parsed_file, file_bytes) = parser
        .parse_file_with_bytes(&path)
        .expect("parse_file_with_bytes");
    assert!(parsed_file.document().is_some());
    assert!(!file_bytes.is_empty());

    let reader = std::fs::File::open(&path).expect("open temp docx");
    let (parsed_reader, reader_bytes) = parser
        .parse_reader_with_bytes(reader)
        .expect("parse_reader_with_bytes");
    assert!(parsed_reader.document().is_some());
    assert_eq!(file_bytes, reader_bytes);

    std::fs::remove_file(path).ok();
}
