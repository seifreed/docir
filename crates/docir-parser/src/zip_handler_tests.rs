use super::zip_handler_validation::is_path_traversal;
use super::*;
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

fn make_zip(entries: &[(&str, &[u8])], method: CompressionMethod) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(cursor);
    let options = SimpleFileOptions::default().compression_method(method);

    for (name, contents) in entries {
        writer.start_file(name, options).expect("start file");
        writer.write_all(contents).expect("write file");
    }

    writer.finish().expect("finish zip").into_inner()
}

fn make_duplicate_name_zip() -> Vec<u8> {
    fn push_u16(out: &mut Vec<u8>, value: u16) {
        out.extend_from_slice(&value.to_le_bytes());
    }

    fn push_u32(out: &mut Vec<u8>, value: u32) {
        out.extend_from_slice(&value.to_le_bytes());
    }

    fn push_local_file(out: &mut Vec<u8>, name: &[u8]) -> u32 {
        let offset = out.len() as u32;
        push_u32(out, 0x0403_4b50);
        push_u16(out, 20);
        push_u16(out, 0);
        push_u16(out, 0);
        push_u16(out, 0);
        push_u16(out, 0);
        push_u32(out, 0);
        push_u32(out, 0);
        push_u32(out, 0);
        push_u16(out, name.len() as u16);
        push_u16(out, 0);
        out.extend_from_slice(name);
        offset
    }

    fn push_central_file(out: &mut Vec<u8>, name: &[u8], local_offset: u32) {
        push_u32(out, 0x0201_4b50);
        push_u16(out, 20);
        push_u16(out, 20);
        push_u16(out, 0);
        push_u16(out, 0);
        push_u16(out, 0);
        push_u16(out, 0);
        push_u32(out, 0);
        push_u32(out, 0);
        push_u32(out, 0);
        push_u16(out, name.len() as u16);
        push_u16(out, 0);
        push_u16(out, 0);
        push_u16(out, 0);
        push_u16(out, 0);
        push_u32(out, 0);
        push_u32(out, local_offset);
        out.extend_from_slice(name);
    }

    let name = b"word/document.xml";
    let mut out = Vec::new();
    let first = push_local_file(&mut out, name);
    let second = push_local_file(&mut out, name);
    let central_offset = out.len() as u32;
    push_central_file(&mut out, name, first);
    push_central_file(&mut out, name, second);
    let central_size = out.len() as u32 - central_offset;
    push_u32(&mut out, 0x0605_4b50);
    push_u16(&mut out, 0);
    push_u16(&mut out, 0);
    push_u16(&mut out, 2);
    push_u16(&mut out, 2);
    push_u32(&mut out, central_size);
    push_u32(&mut out, central_offset);
    push_u16(&mut out, 0);
    out
}

#[test]
fn test_path_traversal_detection() {
    assert!(is_path_traversal("../etc/passwd"));
    assert!(is_path_traversal("foo/../bar"));
    assert!(is_path_traversal("/absolute/path"));
    assert!(is_path_traversal("C:\\Windows"));
    assert!(is_path_traversal("C:Windows"));
    assert!(is_path_traversal("foo\\bar"));

    assert!(!is_path_traversal("word/document.xml"));
    assert!(!is_path_traversal("[Content_Types].xml"));
    assert!(!is_path_traversal("_rels/.rels"));
}

#[test]
fn secure_zip_reader_reads_and_lists_files() {
    let bytes = make_zip(
        &[
            ("word/document.xml", b"<doc/>"),
            ("word/_rels/document.xml.rels", b"<rels/>"),
        ],
        CompressionMethod::Stored,
    );

    let mut reader =
        SecureZipReader::new(Cursor::new(bytes), ZipConfig::default()).expect("reader");

    assert!(!reader.is_empty());
    assert_eq!(reader.len(), 2);
    assert!(reader.contains("word/document.xml"));
    assert_eq!(
        reader.read_file("word/document.xml").expect("bytes"),
        b"<doc/>".to_vec()
    );
    assert_eq!(
        reader
            .read_file_string("word/_rels/document.xml.rels")
            .expect("string"),
        "<rels/>"
    );
    assert_eq!(reader.file_size("word/document.xml").expect("size"), 6);

    let mut names = reader
        .file_names()
        .map(ToString::to_string)
        .collect::<Vec<_>>();
    names.sort();
    assert_eq!(
        names,
        vec![
            "word/_rels/document.xml.rels".to_string(),
            "word/document.xml".to_string()
        ]
    );

    let mut prefix = reader.list_prefix("word/").to_vec();
    prefix.sort();
    assert_eq!(prefix.len(), 2);
    assert_eq!(
        reader.list_suffix(".rels"),
        vec!["word/_rels/document.xml.rels"]
    );
}

#[test]
fn secure_zip_reader_reports_missing_and_encoding_errors() {
    let bytes = make_zip(
        &[
            ("word/document.xml", b"<doc/>"),
            ("word/binary.bin", &[0xff, 0xfe]),
        ],
        CompressionMethod::Stored,
    );
    let mut reader =
        SecureZipReader::new(Cursor::new(bytes), ZipConfig::default()).expect("reader");

    let err = reader.read_file("word/missing.xml").unwrap_err();
    assert!(matches!(err, ParseError::MissingPart(_)));

    let err = reader.read_file_string("word/binary.bin").unwrap_err();
    assert!(matches!(err, ParseError::Encoding(_)));
}

#[test]
fn secure_zip_reader_rejects_duplicate_file_names() {
    let bytes = make_duplicate_name_zip();

    let err = SecureZipReader::new(Cursor::new(bytes), ZipConfig::default())
        .err()
        .expect("duplicate file name error");
    assert!(matches!(err, ParseError::InvalidZip(_)));
}

#[test]
fn secure_zip_reader_rejects_archive_level_limits() {
    let bytes = make_zip(
        &[
            ("a.xml", b"123"),
            ("b.xml", b"456"),
            ("c.xml", b"789"),
            ("deep/path/item.xml", b"x"),
        ],
        CompressionMethod::Stored,
    );

    let count_err = SecureZipReader::new(
        Cursor::new(bytes.clone()),
        ZipConfig {
            max_file_count: 2,
            ..ZipConfig::default()
        },
    )
    .err()
    .expect("file count limit error");
    assert!(matches!(count_err, ParseError::ResourceLimit(_)));

    let depth_err = SecureZipReader::new(
        Cursor::new(bytes.clone()),
        ZipConfig {
            max_path_depth: 1,
            ..ZipConfig::default()
        },
    )
    .err()
    .expect("path depth limit error");
    assert!(matches!(depth_err, ParseError::ResourceLimit(_)));

    let total_err = SecureZipReader::new(
        Cursor::new(bytes),
        ZipConfig {
            max_total_size: 3,
            ..ZipConfig::default()
        },
    )
    .err()
    .expect("total size limit error");
    assert!(matches!(total_err, ParseError::ResourceLimit(_)));
}

#[test]
fn secure_zip_reader_rejects_path_traversal_and_large_files() {
    let traversal = make_zip(&[("../evil.xml", b"x")], CompressionMethod::Stored);
    let err = SecureZipReader::new(Cursor::new(traversal), ZipConfig::default())
        .err()
        .expect("path traversal error");
    assert!(matches!(err, ParseError::PathTraversal(_)));

    let bytes = make_zip(
        &[("word/document.xml", b"12345")],
        CompressionMethod::Stored,
    );
    let err = SecureZipReader::new(
        Cursor::new(bytes),
        ZipConfig {
            max_file_size: 4,
            ..ZipConfig::default()
        },
    )
    .err()
    .expect("file size limit error");
    assert!(matches!(err, ParseError::ResourceLimit(_)));
}

#[test]
fn secure_zip_reader_rejects_suspicious_compression_ratio() {
    let large = vec![b'A'; 3 * 1024 * 1024];
    let bytes = make_zip(
        &[("word/document.xml", &large)],
        CompressionMethod::Deflated,
    );

    let err = SecureZipReader::new(
        Cursor::new(bytes),
        ZipConfig {
            max_file_size: 4 * 1024 * 1024,
            max_total_size: 4 * 1024 * 1024,
            max_compression_ratio: 1.0,
            ..ZipConfig::default()
        },
    )
    .err()
    .expect("compression ratio limit error");
    assert!(matches!(err, ParseError::ResourceLimit(_)));
}
