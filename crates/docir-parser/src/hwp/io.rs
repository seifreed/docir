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
    compressed: bool,
    encrypted: bool,
    password: Option<&str>,
    force_parse: bool,
    try_raw_encrypted: bool,
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

            if compressed {
                if let Some(bytes) = prepare_hwp_stream_data(
                    &data,
                    encrypted,
                    password,
                    force_parse,
                    try_raw_encrypted,
                    path,
                    diagnostics,
                ) {
                    match super::maybe_decompress_stream(&bytes, compressed, path) {
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
                path, size, compressed, hash_hex, decompress_status
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
    if data.len() % 16 != 0 {
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
