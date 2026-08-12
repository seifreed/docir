use super::container::detect_odf_format;
use super::container_encryption::decrypt_odf_part;
use super::container_meta::parse_meta;
use super::*;

fn encryption_data() -> OdfEncryptionData {
    OdfEncryptionData {
        checksum_type: None,
        checksum: None,
        algorithm_name: Some("http://www.w3.org/2001/04/xmlenc#aes256-cbc".to_string()),
        init_vector: Some(vec![0_u8; 16]),
        key_derivation_name: None,
        salt: Some(vec![1_u8; 16]),
        iteration_count: Some(10),
        key_size: Some(256),
    }
}

#[test]
fn detect_odf_format_supports_expected_mimetypes() {
    assert_eq!(
        detect_odf_format("application/vnd.oasis.opendocument.text"),
        Some(DocumentFormat::OdfText)
    );
    assert_eq!(
        detect_odf_format("application/vnd.sun.xml.calc"),
        Some(DocumentFormat::OdfSpreadsheet)
    );
    assert_eq!(
        detect_odf_format("application/vnd.oasis.opendocument.presentation"),
        Some(DocumentFormat::OdfPresentation)
    );
    assert_eq!(detect_odf_format("application/octet-stream"), None);
}

#[test]
fn parse_meta_extracts_known_fields_and_handles_empty_or_malformed_xml() {
    let meta = parse_meta(
        r#"
            <office:meta xmlns:dc="dc" xmlns:meta="meta">
              <dc:title>Title</dc:title>
              <dc:subject>Subject</dc:subject>
              <dc:creator>Alice</dc:creator>
              <meta:keyword>tag1,tag2</meta:keyword>
              <dc:description>Desc</dc:description>
              <meta:creation-date>2026-01-01</meta:creation-date>
              <dc:date>2026-01-02</dc:date>
            </office:meta>
            "#,
    )
    .expect("meta must parse");
    let meta = meta.expect("metadata fields must be present");
    assert_eq!(meta.title.as_deref(), Some("Title"));
    assert_eq!(meta.subject.as_deref(), Some("Subject"));
    assert_eq!(meta.creator.as_deref(), Some("Alice"));
    assert_eq!(meta.keywords.as_deref(), Some("tag1,tag2"));
    assert_eq!(meta.description.as_deref(), Some("Desc"));
    assert_eq!(meta.created.as_deref(), Some("2026-01-01"));
    assert_eq!(meta.modified.as_deref(), Some("2026-01-02"));

    let prefixed_meta = parse_meta(
        r#"
            <pkg:meta xmlns:dct="dc" xmlns:m="meta">
              <dct:title>Alt Title</dct:title>
              <dct:creator>Bob</dct:creator>
              <m:creation-date>2026-02-01</m:creation-date>
            </pkg:meta>
            "#,
    )
    .expect("alternate-prefix meta must parse");
    let prefixed_meta = prefixed_meta.expect("metadata fields must be present");
    assert_eq!(prefixed_meta.title.as_deref(), Some("Alt Title"));
    assert_eq!(prefixed_meta.creator.as_deref(), Some("Bob"));
    assert_eq!(prefixed_meta.created.as_deref(), Some("2026-02-01"));

    assert!(
        parse_meta("<office:meta/>")
            .expect("empty metadata must parse")
            .is_none()
    );
    let err = parse_meta("<office:meta><dc:title>").expect_err("malformed metadata must fail");
    assert!(matches!(err, ParseError::Xml { file, .. } if file == "meta.xml"));
}

#[test]
fn decrypt_odf_part_validates_required_encryption_fields() {
    let mut enc = encryption_data();
    enc.salt = None;
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("missing salt");
    assert!(err.contains("Missing encryption salt"));

    let mut enc = encryption_data();
    enc.init_vector = None;
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("missing iv");
    assert!(err.contains("Missing encryption IV"));

    let mut enc = encryption_data();
    enc.init_vector = Some(vec![0_u8; 8]);
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("bad iv length");
    assert!(err.contains("Unsupported IV length: 8"));
}

#[test]
fn decrypt_odf_part_rejects_unsupported_algorithm_or_key_length() {
    let mut enc = encryption_data();
    enc.algorithm_name = Some("urn:unknown".to_string());
    enc.key_size = None;
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("unsupported algo");
    assert!(err.contains("Unsupported encryption algorithm"));

    let mut enc = encryption_data();
    enc.key_size = Some(192);
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("unsupported key size");
    assert!(err.contains("Unsupported key length: 24"));
}
