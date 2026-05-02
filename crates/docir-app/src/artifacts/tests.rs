use super::*;
use crate::artifacts::ole::parse_ole10_native;
use crate::{test_support::build_test_cfb, DocirApp, ParserConfig};
use docir_core::ExtractedArtifactKind;
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;

fn build_minimal_docx(extra_entries: &[(&str, &[u8])]) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::<u8>::new());
    {
        let mut writer = zip::ZipWriter::new(&mut cursor);
        let options = SimpleFileOptions::default();
        writer
            .start_file("[Content_Types].xml", options)
            .expect("content types");
        writer
            .write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
            )
            .expect("write content types");
        writer.start_file("_rels/.rels", options).expect("rels");
        writer
            .write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
            )
            .expect("write rels");
        writer
            .start_file("word/document.xml", options)
            .expect("document");
        writer
            .write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>docir</w:t></w:r></w:p></w:body>
</w:document>"#,
            )
            .expect("write document");
        for (path, data) in extra_entries {
            writer.start_file(path, options).expect("extra entry");
            writer.write_all(data).expect("extra data");
        }
        writer.finish().expect("finish zip");
    }
    cursor.into_inner()
}

#[test]
fn parse_ole10_native_extracts_metadata_and_payload() {
    let mut blob = Vec::new();
    blob.extend_from_slice(&100u32.to_le_bytes());
    blob.extend_from_slice(&2u16.to_le_bytes());
    blob.extend_from_slice(b"payload.exe\0");
    blob.extend_from_slice(b"C:\\src\\payload.exe\0");
    blob.extend_from_slice(&0u32.to_le_bytes());
    blob.extend_from_slice(&0u32.to_le_bytes());
    blob.extend_from_slice(b"C:\\temp\\payload.exe\0");
    blob.extend_from_slice(&4u32.to_le_bytes());
    blob.extend_from_slice(b"MZ!!");

    let payload = parse_ole10_native(&blob).expect("payload");
    assert_eq!(payload.file_name.as_deref(), Some("payload.exe"));
    assert_eq!(payload.source_path.as_deref(), Some("C:\\src\\payload.exe"));
    assert_eq!(payload.temp_path.as_deref(), Some("C:\\temp\\payload.exe"));
    assert_eq!(payload.data, b"MZ!!");
}

#[test]
fn scan_rtf_objdata_decodes_embedded_hex() {
    let blobs = scan_rtf_objdata(br"{\rtf1{\object{\objdata 4d5a9000}}}");
    assert_eq!(blobs, vec![vec![0x4d, 0x5a, 0x90, 0x00]]);
}

#[test]
fn extract_artifacts_finds_ooxml_embedding_and_payload() {
    let bytes = build_minimal_docx(&[("word/embeddings/object1.bin", b"MZPAYLOAD")]);
    let app = DocirApp::new(ParserConfig::default());
    let parsed = app.parse_bytes(&bytes).expect("parse bytes");
    let bundle = extract_artifacts_from_bytes(
        &parsed,
        &bytes,
        Some("memory.docx".to_string()),
        &ParserConfig::default().zip_config,
        &ArtifactExtractionOptions {
            with_raw: true,
            ..ArtifactExtractionOptions::default()
        },
    );

    assert!(bundle
        .manifest
        .artifacts
        .iter()
        .any(|artifact| artifact.kind == ExtractedArtifactKind::OleObject));
    assert!(bundle
        .payloads
        .iter()
        .any(|payload| payload.relative_path.starts_with("raw/")));
}

#[test]
fn extract_artifacts_finds_legacy_cfb_payload() {
    let mut ole10 = Vec::new();
    ole10.extend_from_slice(&64u32.to_le_bytes());
    ole10.extend_from_slice(&2u16.to_le_bytes());
    ole10.extend_from_slice(b"dropper.exe\0");
    ole10.extend_from_slice(b"C:\\src\\dropper.exe\0");
    ole10.extend_from_slice(&0u32.to_le_bytes());
    ole10.extend_from_slice(&0u32.to_le_bytes());
    ole10.extend_from_slice(b"C:\\temp\\dropper.exe\0");
    ole10.extend_from_slice(&4u32.to_le_bytes());
    ole10.extend_from_slice(b"MZ!!");

    let bytes = build_test_cfb(&[("WordDocument", b"doc"), ("Ole10Native", &ole10)]);
    let app = DocirApp::new(ParserConfig::default());
    let parsed = app.parse_bytes(&bytes).expect("parse legacy doc");
    let bundle = extract_artifacts_from_bytes(
        &parsed,
        &bytes,
        Some("legacy.doc".to_string()),
        &ParserConfig::default().zip_config,
        &ArtifactExtractionOptions::default(),
    );

    assert!(bundle
        .manifest
        .artifacts
        .iter()
        .any(|artifact| artifact.kind == ExtractedArtifactKind::OleObject));
    assert!(bundle
        .manifest
        .artifacts
        .iter()
        .any(|artifact| artifact.kind == ExtractedArtifactKind::OleObject
            && artifact.start_sector.is_some()));
    assert!(bundle
        .manifest
        .artifacts
        .iter()
        .any(|artifact| artifact.kind == ExtractedArtifactKind::OleNativePayload));
}

#[test]
fn extract_artifacts_finds_legacy_package_payload() {
    let bytes = build_test_cfb(&[("Workbook", b"wb"), ("Package", b"%PDF-1.7")]);
    let app = DocirApp::new(ParserConfig::default());
    let parsed = app.parse_bytes(&bytes).expect("parse legacy xls");
    let bundle = extract_artifacts_from_bytes(
        &parsed,
        &bytes,
        Some("legacy.xls".to_string()),
        &ParserConfig::default().zip_config,
        &ArtifactExtractionOptions::default(),
    );

    assert!(bundle.manifest.artifacts.iter().any(|artifact| {
        artifact.kind == ExtractedArtifactKind::OleObject
            && artifact.source_path.as_deref() == Some("Package")
    }));
    assert!(bundle.manifest.artifacts.iter().any(|artifact| {
        artifact.kind == ExtractedArtifactKind::OleNativePayload
            && artifact.mime_type.as_deref() == Some("application/pdf")
    }));
}

#[test]
fn extract_artifacts_finds_legacy_ppt_package_payload() {
    let bytes = build_test_cfb(&[
        ("PowerPoint Document", b"ppt"),
        ("Current User", b"user"),
        ("Package", b"%PDF-1.7"),
    ]);
    let app = DocirApp::new(ParserConfig::default());
    let parsed = app.parse_bytes(&bytes).expect("parse legacy ppt");
    let bundle = extract_artifacts_from_bytes(
        &parsed,
        &bytes,
        Some("legacy.ppt".to_string()),
        &ParserConfig::default().zip_config,
        &ArtifactExtractionOptions::default(),
    );

    assert!(bundle.manifest.artifacts.iter().any(|artifact| {
        artifact.kind == ExtractedArtifactKind::OleObject
            && artifact.source_path.as_deref() == Some("Package")
    }));
    assert!(bundle.manifest.artifacts.iter().any(|artifact| {
        artifact.kind == ExtractedArtifactKind::OleNativePayload
            && artifact.mime_type.as_deref() == Some("application/pdf")
    }));
}

#[test]
fn extract_artifacts_finds_objectpool_package_payload() {
    let bytes = build_test_cfb(&[
        ("WordDocument", b"doc"),
        ("ObjectPool/1/Package", b"%PDF-1.7"),
    ]);
    let app = DocirApp::new(ParserConfig::default());
    let parsed = app.parse_bytes(&bytes).expect("parse objectpool package");
    let bundle = extract_artifacts_from_bytes(
        &parsed,
        &bytes,
        Some("objectpool.doc".to_string()),
        &ParserConfig::default().zip_config,
        &ArtifactExtractionOptions::default(),
    );

    assert!(bundle.manifest.artifacts.iter().any(|artifact| {
        artifact.kind == ExtractedArtifactKind::OleObject
            && artifact.source_path.as_deref() == Some("ObjectPool/1/Package")
    }));
    assert!(bundle.manifest.artifacts.iter().any(|artifact| {
        artifact.kind == ExtractedArtifactKind::OleNativePayload
            && artifact.source_path.as_deref() == Some("ObjectPool/1/Package#ObjectPool/1/Package")
    }));
}

#[test]
fn extract_artifacts_finds_objectpool_ole10native_payload() {
    let mut ole10 = Vec::new();
    ole10.extend_from_slice(&64u32.to_le_bytes());
    ole10.extend_from_slice(&2u16.to_le_bytes());
    ole10.extend_from_slice(b"dropper.exe\0");
    ole10.extend_from_slice(b"C:\\src\\dropper.exe\0");
    ole10.extend_from_slice(&0u32.to_le_bytes());
    ole10.extend_from_slice(&0u32.to_le_bytes());
    ole10.extend_from_slice(b"C:\\temp\\dropper.exe\0");
    ole10.extend_from_slice(&4u32.to_le_bytes());
    ole10.extend_from_slice(b"MZ!!");

    let bytes = build_test_cfb(&[
        ("WordDocument", b"doc"),
        ("ObjectPool/1/Ole10Native", &ole10),
    ]);
    let app = DocirApp::new(ParserConfig::default());
    let parsed = app
        .parse_bytes(&bytes)
        .expect("parse objectpool ole10native");
    let bundle = extract_artifacts_from_bytes(
        &parsed,
        &bytes,
        Some("objectpool.doc".to_string()),
        &ParserConfig::default().zip_config,
        &ArtifactExtractionOptions::default(),
    );

    assert!(bundle.manifest.artifacts.iter().any(|artifact| {
        artifact.kind == ExtractedArtifactKind::OleObject
            && artifact.source_path.as_deref() == Some("ObjectPool/1/Ole10Native")
    }));
    assert!(bundle.manifest.artifacts.iter().any(|artifact| {
        artifact.kind == ExtractedArtifactKind::OleNativePayload
            && artifact.suggested_name.as_deref() == Some("dropper.exe")
    }));
}
