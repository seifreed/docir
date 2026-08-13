use aes::{Aes128, Aes256};
use cbc::Decryptor;
use cbc::cipher::{BlockDecryptMut, KeyIvInit, block_padding::Pkcs7};
use pbkdf2::pbkdf2_hmac;
use sha1::Sha1;

use super::OdfEncryptionData;

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
    let key_bits = encryption
        .key_size
        .or_else(|| {
            if algorithm.contains("aes256") {
                Some(256)
            } else if algorithm.contains("aes128") {
                Some(128)
            } else {
                None
            }
        })
        .ok_or_else(|| "Unsupported encryption algorithm".to_string())?;
    let key_len = (key_bits / 8) as usize;
    if iv.len() != 16 {
        return Err(format!("Unsupported IV length: {}", iv.len()));
    }

    let mut key = vec![0u8; key_len];
    // Security note: PBKDF2 with SHA-1 is required by the ODF specification
    // (OpenDocument 1.3, section 4.4). SHA-1 is deprecated for cryptographic
    // use but must be used here for spec compliance.
    pbkdf2_hmac::<Sha1>(password.as_bytes(), salt, iterations, &mut key);

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
