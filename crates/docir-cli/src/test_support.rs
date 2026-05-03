//! Shared test utilities for CLI command tests.

use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};
use std::{fs, io::Write};
use zip::write::SimpleFileOptions;

/// Path to a named fixture file in the ooxml fixtures directory.
pub fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/ooxml")
        .join(name)
}

/// Create a temporary file path with a custom extension.
pub fn temp_file(name: &str, ext: &str) -> PathBuf {
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("clock")
        .as_nanos();
    std::env::temp_dir().join(format!("docir_cli_{name}_{nanos}.{ext}"))
}

/// Create a temporary directory path.
pub fn temp_dir(name: &str) -> PathBuf {
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("clock")
        .as_nanos();
    std::env::temp_dir().join(format!("docir_cli_{name}_{nanos}"))
}

/// Create a minimal DOCX file at the given path.
pub fn write_docx(path: &PathBuf) {
    let file = fs::File::create(path).expect("create docx");
    let mut zip = zip::ZipWriter::new(file);
    let options = SimpleFileOptions::default();
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
