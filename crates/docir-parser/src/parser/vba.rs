use docir_core::security::{MacroModuleType, MacroReference};

type VbaProjectTextParse = (
    Option<String>,
    Vec<(String, MacroModuleType)>,
    Vec<MacroReference>,
    bool,
);

const MAX_VBA_DECOMPRESSED_SIZE: usize = 10 * 1024 * 1024;
const MAX_VBA_CHUNK_SIZE: usize = 4096;

pub(super) fn parse_vba_project_text(text: &str) -> VbaProjectTextParse {
    let mut project_name = None;
    let mut modules = Vec::new();
    let mut references = Vec::new();
    let mut protected = false;

    for line in text.lines() {
        let line = line.trim();
        if line.is_empty() {
            continue;
        }
        if line.starts_with("Name=") {
            project_name = Some(
                line.strip_prefix("Name=")
                    .unwrap_or(line)
                    .trim()
                    .trim_matches('"')
                    .to_string(),
            );
        } else if line.starts_with("Module=") {
            let name = line
                .strip_prefix("Module=")
                .unwrap_or(line)
                .split('/')
                .next()
                .unwrap_or("")
                .to_string();
            if !name.is_empty() {
                modules.push((name, MacroModuleType::Standard));
            }
        } else if line.starts_with("Class=") {
            let name = line
                .strip_prefix("Class=")
                .unwrap_or(line)
                .split('/')
                .next()
                .unwrap_or("")
                .to_string();
            if !name.is_empty() {
                modules.push((name, MacroModuleType::Class));
            }
        } else if line.starts_with("Document=") {
            let name = line
                .strip_prefix("Document=")
                .unwrap_or(line)
                .split('/')
                .next()
                .unwrap_or("")
                .to_string();
            if !name.is_empty() {
                modules.push((name, MacroModuleType::Document));
            }
        } else if line.starts_with("Reference=") {
            references.push(MacroReference {
                name: line.to_string(),
                guid: None,
                path: None,
                major_version: None,
                minor_version: None,
            });
        } else if line.starts_with("DPB=") {
            protected = true;
        }
    }

    (project_name, modules, references, protected)
}

pub(super) fn vba_decompress(data: &[u8]) -> Option<Vec<u8>> {
    if data.is_empty() {
        return None;
    }
    if data[0] != 0x01 {
        if data.len() > MAX_VBA_DECOMPRESSED_SIZE {
            return None;
        }
        return Some(data.to_vec());
    }

    let mut out = Vec::new();
    let mut pos = 1usize;
    let mut saw_chunk = false;
    while pos + 2 <= data.len() {
        let chunk_start = pos;
        let header = u16::from_le_bytes([data[pos], data[pos + 1]]);
        pos += 2;
        let chunk_size = ((header & 0x0FFF) as usize) + 3;
        let compressed = (header & 0x8000) != 0;
        if (header >> 12) & 0x07 != 0x03
            || (!compressed && chunk_size != MAX_VBA_CHUNK_SIZE + 2)
            || (compressed && chunk_size > MAX_VBA_CHUNK_SIZE + 2)
        {
            return None;
        }
        let chunk_end = chunk_start.checked_add(chunk_size)?;
        if chunk_end > data.len() {
            return None;
        }
        saw_chunk = true;
        if !compressed {
            if out.len().saturating_add(MAX_VBA_CHUNK_SIZE) > MAX_VBA_DECOMPRESSED_SIZE {
                return None;
            }
            out.extend_from_slice(&data[pos..chunk_end]);
            pos = chunk_end;
            continue;
        }

        let mut chunk_out = Vec::new();
        while pos < chunk_end {
            let flags = data[pos];
            pos += 1;
            for bit in 0..8 {
                if pos >= chunk_end {
                    break;
                }
                if (flags & (1 << bit)) == 0 {
                    chunk_out.push(data[pos]);
                    pos += 1;
                } else {
                    if pos + 2 > chunk_end {
                        return None;
                    }
                    let token = u16::from_le_bytes([data[pos], data[pos + 1]]);
                    pos += 2;
                    let (offset, length) = decode_copy_token(token, chunk_out.len())?;
                    if chunk_out.len().checked_add(length)? > MAX_VBA_CHUNK_SIZE {
                        return None;
                    }
                    for _ in 0..length {
                        // offset == 0 indicates malformed token (would cause underflow)
                        // offset > chunk_out.len() would cause index out of bounds
                        // decode_copy_token ensures offset >= 1, so this check is defensive
                        if offset == 0 || offset > chunk_out.len() {
                            return None;
                        }
                        let b = chunk_out[chunk_out.len() - offset];
                        chunk_out.push(b);
                    }
                }
            }
        }
        if out.len().saturating_add(chunk_out.len()) > MAX_VBA_DECOMPRESSED_SIZE {
            return None;
        }
        out.extend_from_slice(&chunk_out);
    }

    if !saw_chunk || pos != data.len() {
        return None;
    }
    Some(out)
}

pub(super) fn normalize_vba_source_text(data: &[u8]) -> String {
    let text = String::from_utf8_lossy(data);
    let text = text.strip_prefix('\u{feff}').unwrap_or(&text);
    text.replace("\r\n", "\n").replace('\r', "\n")
}

fn decode_copy_token(token: u16, decompressed_len: usize) -> Option<(usize, usize)> {
    if decompressed_len == 0 || decompressed_len > MAX_VBA_CHUNK_SIZE {
        return None;
    }
    let bit_count = if decompressed_len <= 1 {
        0
    } else {
        (usize::BITS - (decompressed_len - 1).leading_zeros()) as usize
    };
    let offset_bits = bit_count.clamp(4, 12);
    let length_bits = 16 - offset_bits;
    let offset_mask = (1u16 << offset_bits) - 1;
    let length_mask = (1u16 << length_bits) - 1;
    let offset = ((token >> length_bits) & offset_mask) as usize + 1;
    let length = (token & length_mask) as usize + 3;
    Some((offset, length))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_vba_project_text_extracts_modules_refs_and_protection() {
        let project = r#"
            Name="InvoiceMacros"
            Module=Core/Module1
            Class=ThisClass/0
            Document=ThisDocument/&H00000000
            Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation
            DPB="AAAAAA"
        "#;

        let (name, modules, refs, protected) = parse_vba_project_text(project);

        assert_eq!(name.as_deref(), Some("InvoiceMacros"));
        assert_eq!(modules.len(), 3);
        assert_eq!(modules[0], ("Core".to_string(), MacroModuleType::Standard));
        assert_eq!(
            modules[1],
            ("ThisClass".to_string(), MacroModuleType::Class)
        );
        assert_eq!(
            modules[2],
            ("ThisDocument".to_string(), MacroModuleType::Document)
        );
        assert_eq!(refs.len(), 1);
        assert!(refs[0].name.starts_with("Reference="));
        assert!(protected);
    }

    #[test]
    fn parse_vba_project_text_does_not_hide_repeated_field_prefixes() {
        let project = r#"
            Name=Name="Hidden"
            Module=Module=Hidden/0
            Class=Class=HiddenClass/0
            Document=Document=HiddenDoc/0
        "#;

        let (name, modules, _, _) = parse_vba_project_text(project);

        assert_eq!(name.as_deref(), Some(r#"Name="Hidden"#));
        assert_eq!(
            modules,
            vec![
                ("Module=Hidden".to_string(), MacroModuleType::Standard),
                ("Class=HiddenClass".to_string(), MacroModuleType::Class),
                ("Document=HiddenDoc".to_string(), MacroModuleType::Document),
            ]
        );
    }

    #[test]
    fn vba_decompress_handles_plain_payload_and_invalid_header() {
        assert_eq!(vba_decompress(&[]), None);

        let plain = b"not-compressed";
        assert_eq!(vba_decompress(plain), Some(plain.to_vec()));
    }

    #[test]
    fn vba_decompress_handles_literal_and_copy_tokens() {
        // 0x01 signature + one compressed chunk:
        // flags=0b00000100 => literal 'A', literal 'B', then copy token(offset=2, len=3)
        let encoded = [0x01, 0x04, 0xB0, 0x04, b'A', b'B', 0x00, 0x10];
        let out = vba_decompress(&encoded).expect("decompress should succeed");
        assert_eq!(out, b"ABABA");
    }

    #[test]
    fn vba_decompress_advances_by_encoded_chunk_size() {
        let encoded = [0x01, 0x01, 0xB0, 0x00, b'A', 0x01, 0xB0, 0x00, b'B'];

        assert_eq!(vba_decompress(&encoded), Some(b"AB".to_vec()));
    }

    #[test]
    fn vba_decompress_rejects_truncated_chunk() {
        // Declares an uncompressed chunk larger than remaining bytes.
        let encoded = [0x01, 0x10, 0x30, b'A', b'B'];
        assert_eq!(vba_decompress(&encoded), None);
    }

    #[test]
    fn vba_decompress_rejects_invalid_copy_token() {
        // A copy token cannot reference bytes before the start of its chunk.
        let encoded = [0x01, 0x02, 0xB0, 0x01, 0x00, 0x00];
        assert_eq!(vba_decompress(&encoded), None);
    }

    #[test]
    fn vba_decompress_rejects_truncated_copy_token() {
        // The flags byte requests a two-byte copy token, but only one byte remains.
        let encoded = [0x01, 0x01, 0xB0, 0x01, 0x00];
        assert_eq!(vba_decompress(&encoded), None);
    }

    #[test]
    fn vba_decompress_rejects_oversized_plain_payload() {
        let payload = vec![b'A'; MAX_VBA_DECOMPRESSED_SIZE + 1];

        assert_eq!(vba_decompress(&payload), None);
    }

    #[test]
    fn decode_copy_token_uses_spec_chunk_boundaries() {
        assert_eq!(decode_copy_token(0xF000, 16), Some((16, 3)));
        assert_eq!(decode_copy_token(0x7800, 17), Some((16, 3)));
    }

    #[test]
    fn vba_decompress_rejects_copy_expansion_past_chunk_limit() {
        let mut chunk = vec![0x00, b'A', b'B', b'C', b'D', b'E', b'F', b'G', b'H'];
        const COPY_GROUP: [u8; 17] = [
            0xFF, 0x0F, 0x00, 0x0F, 0x00, 0x0F, 0x00, 0x0F, 0x00, 0x0F, 0x00, 0x0F, 0x00, 0x0F,
            0x00, 0x0F, 0x00,
        ];
        for _ in 0..40 {
            chunk.extend_from_slice(&COPY_GROUP);
        }
        let chunk_size = u16::try_from(chunk.len() - 1).expect("test chunk size");
        let mut encoded = vec![0x01];
        encoded.extend_from_slice(&(0xB000 | chunk_size).to_le_bytes());
        encoded.extend_from_slice(&chunk);

        assert_eq!(vba_decompress(&encoded), None);
    }
}
