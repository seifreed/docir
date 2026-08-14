use aes::{Aes128, Aes256};
use cbc::Decryptor;
use cbc::cipher::{BlockDecryptMut, KeyIvInit, block_padding::Pkcs7};
use pbkdf2::pbkdf2_hmac;
use sha1::{Digest as Sha1Digest, Sha1};
use sha2::Sha256;

use super::OdfEncryptionData;

pub(super) const MAX_ODF_PBKDF2_ITERATIONS: u32 = 10_000_000;

pub(super) fn decrypt_odf_part(
    encrypted: Vec<u8>,
    encryption: &OdfEncryptionData,
    password: &str,
) -> Result<Vec<u8>, String> {
    let algorithm = encryption
        .algorithm_name
        .as_deref()
        .ok_or_else(|| "Missing encryption algorithm".to_string())?;
    let salt = encryption
        .salt
        .as_ref()
        .ok_or_else(|| "Missing encryption salt".to_string())?;
    let iv = encryption
        .init_vector
        .as_ref()
        .ok_or_else(|| "Missing encryption IV".to_string())?;
    let key_derivation = encryption
        .key_derivation_name
        .as_deref()
        .ok_or_else(|| "Missing key derivation algorithm".to_string())?;
    if key_derivation != "PBKDF2"
        && key_derivation != "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#pbkdf2"
    {
        return Err(format!(
            "Unsupported key derivation algorithm: {key_derivation}"
        ));
    }
    let iterations = encryption
        .iteration_count
        .ok_or_else(|| "Missing encryption iteration count".to_string())?;
    if iterations == 0 {
        return Err("Invalid encryption iteration count: 0".to_string());
    }
    if iterations > MAX_ODF_PBKDF2_ITERATIONS {
        return Err(format!(
            "ODF encryption iteration count exceeds maximum: {} (max: {})",
            iterations, MAX_ODF_PBKDF2_ITERATIONS
        ));
    }
    let start_key = derive_start_key(
        password,
        encryption.start_key_generation_name.as_deref(),
        encryption.start_key_size,
    )?;
    let expected_key_len = match algorithm {
        "http://www.w3.org/2001/04/xmlenc#aes256-cbc" => 32,
        "http://www.w3.org/2001/04/xmlenc#aes128-cbc" => 16,
        _ => return Err("Unsupported encryption algorithm".to_string()),
    };
    if let Some(key_size) = encryption.key_size
        && key_size != expected_key_len
    {
        return Err(format!("Unsupported key length: {key_size}"));
    }
    let key_len = expected_key_len as usize;
    if iv.len() != 16 {
        return Err(format!("Unsupported IV length: {}", iv.len()));
    }

    let mut key = vec![0u8; key_len];
    // Security note: PBKDF2 with SHA-1 is required by the ODF specification
    // (OpenDocument 1.3, section 4.4). SHA-1 is deprecated for cryptographic
    // use but must be used here for spec compliance.
    pbkdf2_hmac::<Sha1>(&start_key, salt, iterations, &mut key);

    let mut buffer = encrypted;
    if key_len == 32 {
        let decryptor = Decryptor::<Aes256>::new_from_slices(&key, iv)
            .map_err(|_| "Invalid AES-256 key or IV".to_string())?;
        let decrypted = decryptor
            .decrypt_padded_mut::<Pkcs7>(&mut buffer)
            .map_err(|_| "Invalid AES-256 padding".to_string())?;
        Ok(decrypted.to_vec())
    } else if key_len == 16 {
        let decryptor = Decryptor::<Aes128>::new_from_slices(&key, iv)
            .map_err(|_| "Invalid AES-128 key or IV".to_string())?;
        let decrypted = decryptor
            .decrypt_padded_mut::<Pkcs7>(&mut buffer)
            .map_err(|_| "Invalid AES-128 padding".to_string())?;
        Ok(decrypted.to_vec())
    } else {
        Err(format!("Unsupported key length: {}", key_len))
    }
}

fn derive_start_key(
    password: &str,
    algorithm: Option<&str>,
    key_size: Option<u32>,
) -> Result<Vec<u8>, String> {
    let algorithm = algorithm.unwrap_or("SHA1");
    let (expected_size, key) = if algorithm.eq_ignore_ascii_case("SHA1")
        || algorithm.eq_ignore_ascii_case("http://www.w3.org/2000/09/xmldsig#sha1")
    {
        (20, Sha1::digest(password.as_bytes()).to_vec())
    } else if algorithm.eq_ignore_ascii_case("http://www.w3.org/2000/09/xmldsig#sha256")
        || algorithm.eq_ignore_ascii_case("http://www.w3.org/2001/04/xmlenc#sha256")
    {
        (32, Sha256::digest(password.as_bytes()).to_vec())
    } else {
        return Err(format!(
            "Unsupported start key generation algorithm: {algorithm}"
        ));
    };
    if let Some(key_size) = key_size
        && key_size != expected_size
    {
        return Err(format!(
            "Unsupported start key length: {key_size} (expected {expected_size})"
        ));
    }
    Ok(key)
}
