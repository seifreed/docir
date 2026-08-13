use super::container::detect_odf_format;
use super::container_encryption::decrypt_odf_part;
use super::container_meta::parse_meta;
use super::*;
use aes::Aes256;
use cbc::Encryptor;
use cbc::cipher::{BlockEncryptMut, KeyIvInit, block_padding::Pkcs7};
use pbkdf2::pbkdf2_hmac;
use sha1::Sha1;
use sha2::{Digest, Sha256};

fn encryption_data() -> OdfEncryptionData {
    OdfEncryptionData {
        checksum_type: None,
        checksum: None,
        algorithm_name: Some("http://www.w3.org/2001/04/xmlenc#aes256-cbc".to_string()),
        init_vector: Some(vec![0_u8; 16]),
        key_derivation_name: Some("PBKDF2".to_string()),
        salt: Some(vec![1_u8; 16]),
        iteration_count: Some(10),
        key_size: Some(32),
        start_key_generation_name: None,
        start_key_size: None,
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
    enc.algorithm_name = None;
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("missing algorithm");
    assert!(err.contains("Missing encryption algorithm"));

    let mut enc = encryption_data();
    enc.key_derivation_name = None;
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("missing key derivation");
    assert!(err.contains("Missing key derivation algorithm"));

    let mut enc = encryption_data();
    enc.iteration_count = None;
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("missing iterations");
    assert!(err.contains("Missing encryption iteration count"));

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
    enc.key_size = Some(24);
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("unsupported key size");
    assert!(err.contains("Unsupported key length: 24"));

    let mut enc = encryption_data();
    enc.algorithm_name = Some("urn:unknown".to_string());
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw")
        .expect_err("unknown algorithm must not be selected by key size");
    assert!(err.contains("Unsupported encryption algorithm"));

    let mut enc = encryption_data();
    enc.start_key_generation_name = Some("urn:unknown".to_string());
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw")
        .expect_err("unknown start-key algorithm must be rejected");
    assert!(err.contains("Unsupported start key generation algorithm"));

    let mut enc = encryption_data();
    enc.key_size = Some(u32::MAX);
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw")
        .expect_err("unrepresentable key size must be rejected");
    assert!(err.contains("Unsupported key length"));

    let mut enc = encryption_data();
    enc.key_derivation_name = Some("urn:unsupported:kdf".to_string());
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("unsupported kdf");
    assert!(err.contains("Unsupported key derivation algorithm"));

    let mut enc = encryption_data();
    enc.iteration_count = Some(0);
    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw").expect_err("zero iterations");
    assert!(err.contains("Invalid encryption iteration count: 0"));
}

#[test]
fn decrypt_odf_part_rejects_excessive_iteration_count() {
    let mut enc = encryption_data();
    enc.iteration_count = Some(super::container_encryption::MAX_ODF_PBKDF2_ITERATIONS + 1);

    let err = decrypt_odf_part(vec![0_u8; 16], &enc, "pw")
        .expect_err("excessive PBKDF2 iterations must be rejected");
    assert!(err.contains("iteration count exceeds maximum"));
}

#[test]
fn decrypt_odf_part_uses_sha256_start_key_generation() {
    let mut enc = encryption_data();
    enc.start_key_generation_name = Some("http://www.w3.org/2000/09/xmldsig#sha256".to_string());
    enc.start_key_size = Some(32);
    let start_key = Sha256::digest(b"pw");
    let mut key = [0_u8; 32];
    pbkdf2_hmac::<Sha1>(&start_key, &[1_u8; 16], 10, &mut key);
    let plaintext = b"ODF encrypted content";
    let mut encrypted = vec![0_u8; plaintext.len() + 16];
    encrypted[..plaintext.len()].copy_from_slice(plaintext);
    let encrypted = Encryptor::<Aes256>::new_from_slices(&key, &[0_u8; 16])
        .expect("AES-256 parameters")
        .encrypt_padded_mut::<Pkcs7>(&mut encrypted, plaintext.len())
        .expect("PKCS#7 padding")
        .to_vec();

    let decrypted = decrypt_odf_part(encrypted, &enc, "pw").expect("ODF decrypt");
    assert_eq!(decrypted, plaintext);
}
