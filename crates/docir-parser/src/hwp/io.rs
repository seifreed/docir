use crate::diagnostics::{push_info, push_warning};
use crate::error::ParseError;
use crate::ole::Cfb;
use docir_core::ir::Diagnostics;
use sha1::{Digest as Sha1Digest, Sha1};
use sha2::Sha256;

use aes::Aes128;
use cbc::cipher::{block_padding::Pkcs7, BlockDecryptMut, KeyIvInit};
use cbc::Decryptor;

pub(super) fn prepare_hwp_stream_data(
    data: &[u8],
    encrypted: bool,
    password: Option<&str>,
    force_parse: bool,
    try_raw_encrypted: bool,
    source: &str,
    diagnostics: &mut Diagnostics,
) -> Option<Vec<u8>> {
    if !encrypted {
        return Some(data.to_vec());
    }
    if let Some(password) = password {
        match decrypt_hwp_stream(data, password, source) {
            Ok(bytes) => return Some(bytes),
            Err(err) => {
                push_warning(
                    diagnostics,
                    "HWP_DECRYPT_FAIL",
                    err.to_string(),
                    Some(source),
                );
            }
        }
    }
    if force_parse {
        push_warning(
            diagnostics,
            "HWP_FORCE_PARSE_STREAM",
            "HWP force-parse: using raw encrypted stream bytes".to_string(),
            Some(source),
        );
        return Some(data.to_vec());
    }
    if try_raw_encrypted {
        push_warning(
            diagnostics,
            "HWP_ENCRYPTED_RAW_STREAM",
            "HWP encrypted without password: trying raw stream bytes".to_string(),
            Some(source),
        );
        return Some(data.to_vec());
    }
    None
}

pub(super) fn dump_hwp_streams(
    cfb: &Cfb,
    stream_names: &[String],
    header_ctx: &super::builder::HwpHeaderContext<'_>,
    diagnostics: &mut Diagnostics,
) {
    for path in stream_names {
        let size = cfb.stream_size(path).unwrap_or(0);
        let mut sha = Sha256::new();
        let mut hash_hex = "missing".to_string();
        let mut decompress_status = "skip".to_string();
        if let Some(data) = cfb.read_stream(path) {
            sha.update(&data);
            let hash = sha.finalize();
            let mut out = String::with_capacity(hash.len() * 2);
            for byte in hash {
                out.push_str(&format!("{:02x}", byte));
            }
            hash_hex = out;

            if header_ctx.compressed {
                if let Some(bytes) = prepare_hwp_stream_data(
                    &data,
                    header_ctx.encrypted,
                    header_ctx.hwp_password,
                    header_ctx.force_parse,
                    header_ctx.try_raw_encrypted,
                    path,
                    diagnostics,
                ) {
                    match super::maybe_decompress_stream(&bytes, header_ctx.compressed, path) {
                        Ok(_) => decompress_status = "ok".to_string(),
                        Err(err) => {
                            decompress_status = format!("fail: {}", err);
                        }
                    }
                } else {
                    decompress_status = "encrypted".to_string();
                }
            }
        }
        push_info(
            diagnostics,
            "HWP_STREAM_DUMP",
            format!(
                "stream: {}, size={}, compressed={}, sha256={}, decompress={}",
                path, size, header_ctx.compressed, hash_hex, decompress_status
            ),
            Some(path),
        );
    }
}

fn derive_hwp_key(password: &str) -> [u8; 16] {
    let mut hasher = Sha1::new();
    hasher.update(password.as_bytes());
    let digest = hasher.finalize();
    let mut key = [0u8; 16];
    key.copy_from_slice(&digest[..16]);
    key
}

fn decrypt_hwp_stream(data: &[u8], password: &str, source: &str) -> Result<Vec<u8>, ParseError> {
    if !data.len().is_multiple_of(16) {
        return Err(ParseError::InvalidStructure(format!(
            "HWP encrypted stream {} length not aligned",
            source
        )));
    }
    let key = derive_hwp_key(password);
    let iv = [0u8; 16];
    let mut buffer = data.to_vec();
    let decryptor = Decryptor::<Aes128>::new_from_slices(&key, &iv).map_err(|e| {
        ParseError::InvalidStructure(format!(
            "Failed to init HWP decryptor for {}: {}",
            source, e
        ))
    })?;
    let decrypted = decryptor
        .decrypt_padded_mut::<Pkcs7>(&mut buffer)
        .map_err(|e| {
            ParseError::InvalidStructure(format!("Failed to decrypt HWP stream {}: {}", source, e))
        })?;
    Ok(decrypted.to_vec())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use std::path::PathBuf;

    fn load_minimal_hwp_cfb() -> Cfb {
        let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        path.push("../../fixtures/hwp/minimal.hwp");
        let data = fs::read(&path).unwrap_or_else(|e| panic!("failed to read {path:?}: {e}"));
        Cfb::parse(data).expect("parse minimal.hwp as CFB")
    }

    #[test]
    fn derive_hwp_key_is_stable_and_password_sensitive() {
        let a = derive_hwp_key("secret");
        let b = derive_hwp_key("secret");
        let c = derive_hwp_key("different");
        assert_eq!(a, b);
        assert_ne!(a, c);
    }

    #[test]
    fn decrypt_hwp_stream_rejects_non_aligned_input() {
        let err =
            decrypt_hwp_stream(b"1234567890", "pw", "BodyText/Section0").expect_err("must fail");
        assert!(matches!(err, ParseError::InvalidStructure(_)));
    }

    #[test]
    fn prepare_hwp_stream_data_handles_encryption_modes() {
        let mut diagnostics = Diagnostics::new();
        let payload = b"plain-data";

        let plain = prepare_hwp_stream_data(
            payload,
            false,
            None,
            false,
            false,
            "BodyText/Section0",
            &mut diagnostics,
        )
        .expect("unencrypted data");
        assert_eq!(plain, payload);

        let forced = prepare_hwp_stream_data(
            payload,
            true,
            None,
            true,
            false,
            "BodyText/Section1",
            &mut diagnostics,
        )
        .expect("forced parse returns raw bytes");
        assert_eq!(forced, payload);

        let raw = prepare_hwp_stream_data(
            payload,
            true,
            None,
            false,
            true,
            "BodyText/Section2",
            &mut diagnostics,
        )
        .expect("raw encrypted fallback");
        assert_eq!(raw, payload);

        let none = prepare_hwp_stream_data(
            payload,
            true,
            None,
            false,
            false,
            "BodyText/Section3",
            &mut diagnostics,
        );
        assert!(none.is_none());
    }

    #[test]
    fn prepare_hwp_stream_data_reports_decrypt_failure_with_password() {
        let mut diagnostics = Diagnostics::new();
        let result = prepare_hwp_stream_data(
            b"short",
            true,
            Some("pw"),
            false,
            false,
            "BodyText/SectionX",
            &mut diagnostics,
        );
        assert!(result.is_none());
        assert!(diagnostics
            .entries
            .iter()
            .any(|e| e.code == "HWP_DECRYPT_FAIL"));
    }

    #[test]
    fn prepare_hwp_stream_data_force_parse_fallbacks_after_decrypt_failure() {
        let mut diagnostics = Diagnostics::new();
        let payload = b"unaligned";
        let result = prepare_hwp_stream_data(
            payload,
            true,
            Some("pw"),
            true,
            false,
            "BodyText/SectionY",
            &mut diagnostics,
        )
        .expect("force-parse should keep raw bytes after decrypt failure");

        assert_eq!(result, payload);
        assert!(diagnostics
            .entries
            .iter()
            .any(|e| e.code == "HWP_DECRYPT_FAIL"));
        assert!(diagnostics
            .entries
            .iter()
            .any(|e| e.code == "HWP_FORCE_PARSE_STREAM"));
    }

    #[test]
    fn dump_hwp_streams_emits_diagnostics_for_existing_and_missing_streams() {
        let cfb = load_minimal_hwp_cfb();
        let mut stream_names = cfb.list_streams();
        stream_names.push("Missing/Stream".to_string());
        let header_ctx = super::super::builder::HwpHeaderContext {
            compressed: false,
            encrypted: false,
            force_parse: false,
            hwp_password: None,
            try_raw_encrypted: false,
            allow_parse: true,
        };
        let mut diagnostics = Diagnostics::new();

        dump_hwp_streams(&cfb, &stream_names, &header_ctx, &mut diagnostics);

        let dumps: Vec<_> = diagnostics
            .entries
            .iter()
            .filter(|e| e.code == "HWP_STREAM_DUMP")
            .collect();
        assert!(
            dumps.len() >= stream_names.len(),
            "expected one dump line per stream name"
        );
        assert!(dumps.iter().any(|e| e.message.contains("decompress=skip")));
        assert!(dumps.iter().any(|e| e.message.contains("sha256=missing")));
    }

    #[test]
    fn dump_hwp_streams_marks_encrypted_when_stream_cannot_be_prepared() {
        let cfb = load_minimal_hwp_cfb();
        let stream_names = cfb.list_streams();
        let header_ctx = super::super::builder::HwpHeaderContext {
            compressed: true,
            encrypted: true,
            force_parse: false,
            hwp_password: None,
            try_raw_encrypted: false,
            allow_parse: true,
        };
        let mut diagnostics = Diagnostics::new();

        dump_hwp_streams(&cfb, &stream_names, &header_ctx, &mut diagnostics);

        assert!(diagnostics.entries.iter().any(|e| {
            e.code == "HWP_STREAM_DUMP" && e.message.contains("decompress=encrypted")
        }));
    }

    #[test]
    fn dump_hwp_streams_records_decompress_failure_when_flagged_compressed() {
        let cfb = load_minimal_hwp_cfb();
        let stream_names = cfb.list_streams();
        let header_ctx = super::super::builder::HwpHeaderContext {
            compressed: true,
            encrypted: false,
            force_parse: false,
            hwp_password: None,
            try_raw_encrypted: false,
            allow_parse: true,
        };
        let mut diagnostics = Diagnostics::new();

        dump_hwp_streams(&cfb, &stream_names, &header_ctx, &mut diagnostics);

        assert!(diagnostics
            .entries
            .iter()
            .any(|e| e.code == "HWP_STREAM_DUMP" && e.message.contains("decompress=fail:")));
    }
}
