use crate::diagnostics::{push_info, push_warning};
use crate::error::ParseError;
use crate::ole::Cfb;
use docir_core::ir::Diagnostics;
use sha1::{Digest as Sha1Digest, Sha1};
use sha2::Sha256;

use aes::Aes128;
use aes::cipher::{BlockEncrypt, KeyInit, generic_array::GenericArray};

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
        match cfb.try_read_stream(path) {
            Ok(Some(data)) => {
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
            Ok(None) => {}
            Err(err) => push_warning(
                diagnostics,
                "HWP_STREAM_READ_FAIL",
                err.to_string(),
                Some(path),
            ),
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

const MAX_HWP_PASSWORD_BYTES: usize = 80;

/// Derive the AES-128 key used by HWP 5.x password-protected streams.
///
/// Security note: SHA-1 without salt or iterations is a weak key derivation
/// function. This matches the HWP binary format specification and cannot be
/// changed while still reading existing HWP files. A proper
/// KDF (PBKDF2 with salt, or Argon2id) should be used for new designs.
fn derive_hwp_key(password: &str) -> Result<[u8; 16], ParseError> {
    let password = password.as_bytes();
    if password.len() > MAX_HWP_PASSWORD_BYTES {
        return Err(ParseError::ResourceLimit(format!(
            "HWP password exceeds {} bytes",
            MAX_HWP_PASSWORD_BYTES
        )));
    }
    let mut password_material = [0u8; MAX_HWP_PASSWORD_BYTES * 2];
    for (index, byte) in password.iter().copied().enumerate() {
        let previous = if index == 0 { 236 } else { password[index - 1] };
        password_material[index * 2] = previous.rotate_left(1);
        password_material[index * 2 + 1] = byte;
    }
    let mut hasher = Sha1::new();
    hasher.update(&password_material[..password.len() * 2]);
    let digest = hasher.finalize();
    let mut key = [0u8; 16];
    key.copy_from_slice(&digest[..16]);
    Ok(key)
}

fn decrypt_hwp_stream(data: &[u8], password: &str, source: &str) -> Result<Vec<u8>, ParseError> {
    let key = derive_hwp_key(password)?;
    let cipher = Aes128::new_from_slice(&key).map_err(|e| {
        ParseError::InvalidStructure(format!("Failed to init HWP cipher for {}: {}", source, e))
    })?;
    let mut padded = data.to_vec();
    let remainder = padded.len() % 16;
    if remainder != 0 {
        padded.extend(std::iter::repeat_n((16 - remainder) as u8, 16 - remainder));
    }

    let mut feedback = [0u8; 16];
    let mut output = Vec::with_capacity(padded.len());
    for block in padded.chunks_exact(16) {
        let mut decrypted = [0u8; 16];
        decrypted.copy_from_slice(block);
        for bit_index in 0..128 {
            let mut encrypted_feedback = GenericArray::clone_from_slice(&feedback);
            cipher.encrypt_block(&mut encrypted_feedback);
            let keystream_bit = encrypted_feedback[0] & 0x80;
            let input_bit = (decrypted[bit_index / 8] >> (7 - (bit_index % 8))) & 1;
            let mut index = 1;
            for _ in 0..3 {
                let first = feedback[index];
                feedback[index - 1] = feedback[index - 1].wrapping_shl(1) | (first >> 7);
                let second = feedback[index + 1];
                let third = feedback[index + 2];
                feedback[index] = first.wrapping_shl(1) | (second >> 7);
                feedback[index + 1] = second.wrapping_shl(1) | (third >> 7);
                let fourth = feedback[index + 3];
                feedback[index + 2] = third.wrapping_shl(1) | (fourth >> 7);
                feedback[index + 3] = fourth.wrapping_shl(1) | (feedback[index + 4] >> 7);
                index += 5;
            }
            feedback[15] = feedback[15].wrapping_shl(1) | input_bit;
            decrypted[bit_index / 8] ^= keystream_bit >> (bit_index % 8);
        }
        output.extend_from_slice(&decrypted);
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_support::build_test_cfb;
    use std::fs;
    use std::path::PathBuf;

    fn load_minimal_hwp_cfb() -> Cfb {
        let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        path.push("../../fixtures/hwp/minimal.hwp");
        let data = fs::read(&path).unwrap_or_else(|e| panic!("failed to read {path:?}: {e}"));
        Cfb::parse(data).expect("parse minimal.hwp as CFB")
    }

    fn patch_cfb_fat_entry(bytes: &[u8], fat_index: u32, value: u32) -> Vec<u8> {
        let mut out = bytes.to_vec();
        let sector_size = 1usize << u16::from_le_bytes([out[0x1E], out[0x1F]]);
        let first_fat_sector = u32::from_le_bytes([out[0x4C], out[0x4D], out[0x4E], out[0x4F]]);
        let fat_offset =
            sector_size + first_fat_sector as usize * sector_size + fat_index as usize * 4;
        out[fat_offset..fat_offset + 4].copy_from_slice(&value.to_le_bytes());
        out
    }

    #[test]
    fn derive_hwp_key_is_stable_and_password_sensitive() {
        let key_a = derive_hwp_key("secret").expect("key");
        let key_b = derive_hwp_key("secret").expect("key");
        let key_c = derive_hwp_key("different").expect("key");
        assert_eq!(key_a, key_b);
        assert_ne!(key_a, key_c);
        assert_eq!(
            key_a,
            [
                0x54, 0xd6, 0x91, 0xad, 0x45, 0x96, 0x2e, 0xa8, 0x95, 0xc0, 0xbf, 0xf2, 0x9e, 0xc2,
                0x0d, 0xae,
            ]
        );
    }

    #[test]
    fn decrypt_hwp_stream_matches_hwp_v5_vector() {
        let ciphertext = [
            0xf0, 0x4e, 0x52, 0xaa, 0xc8, 0xa8, 0xe1, 0x04, 0x6f, 0x95, 0x5d, 0x82, 0x07, 0xf1,
            0x88, 0xa9,
        ];

        let plaintext = decrypt_hwp_stream(&ciphertext, "secret", "BodyText/Section0")
            .expect("HWP vector must decrypt");

        assert_eq!(plaintext, b"0123456789ABCDEF");
    }

    #[test]
    fn derive_hwp_key_rejects_oversized_passwords() {
        let password = "x".repeat(MAX_HWP_PASSWORD_BYTES + 1);

        let err = derive_hwp_key(&password).expect_err("oversized password must fail");

        assert!(matches!(err, ParseError::ResourceLimit(_)));
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
        let password = "x".repeat(MAX_HWP_PASSWORD_BYTES + 1);
        let result = prepare_hwp_stream_data(
            b"ciphertext",
            true,
            Some(&password),
            false,
            false,
            "BodyText/SectionX",
            &mut diagnostics,
        );
        assert!(result.is_none());
        assert!(
            diagnostics
                .entries
                .iter()
                .any(|e| e.code == "HWP_DECRYPT_FAIL")
        );
    }

    #[test]
    fn prepare_hwp_stream_data_force_parse_fallbacks_after_decrypt_failure() {
        let mut diagnostics = Diagnostics::new();
        let payload = b"ciphertext";
        let password = "x".repeat(MAX_HWP_PASSWORD_BYTES + 1);
        let result = prepare_hwp_stream_data(
            payload,
            true,
            Some(&password),
            true,
            false,
            "BodyText/SectionY",
            &mut diagnostics,
        )
        .expect("force-parse should keep raw bytes after decrypt failure");

        assert_eq!(result, payload);
        assert!(
            diagnostics
                .entries
                .iter()
                .any(|e| e.code == "HWP_DECRYPT_FAIL")
        );
        assert!(
            diagnostics
                .entries
                .iter()
                .any(|e| e.code == "HWP_FORCE_PARSE_STREAM")
        );
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
    fn dump_hwp_streams_reports_corrupt_stream_read() {
        let mut header = vec![0u8; 5000];
        header[..17].copy_from_slice(b"HWP Document File");
        let base = build_test_cfb(&[("FileHeader", &header)]);
        let bytes = patch_cfb_fat_entry(&base, 0, 99);
        let cfb = Cfb::parse(bytes).expect("cfb");
        let header_ctx = super::super::builder::HwpHeaderContext {
            compressed: false,
            encrypted: false,
            force_parse: false,
            hwp_password: None,
            try_raw_encrypted: false,
            allow_parse: true,
        };
        let mut diagnostics = Diagnostics::new();

        dump_hwp_streams(
            &cfb,
            &["FileHeader".to_string()],
            &header_ctx,
            &mut diagnostics,
        );

        assert!(diagnostics.entries.iter().any(|e| {
            e.code == "HWP_STREAM_READ_FAIL" && e.message.contains("OLE sector out of bounds")
        }));
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

        assert!(
            diagnostics
                .entries
                .iter()
                .any(|e| e.code == "HWP_STREAM_DUMP" && e.message.contains("decompress=fail:"))
        );
    }
}
