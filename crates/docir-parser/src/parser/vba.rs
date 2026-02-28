use docir_core::security::{MacroModuleType, MacroReference};

pub(super) fn parse_vba_project_text(
    text: &str,
) -> (
    Option<String>,
    Vec<(String, MacroModuleType)>,
    Vec<MacroReference>,
    bool,
) {
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
                line.trim_start_matches("Name=")
                    .trim()
                    .trim_matches('"')
                    .to_string(),
            );
        } else if line.starts_with("Module=") {
            let name = line
                .trim_start_matches("Module=")
                .split('/')
                .next()
                .unwrap_or("")
                .to_string();
            if !name.is_empty() {
                modules.push((name, MacroModuleType::Standard));
            }
        } else if line.starts_with("Class=") {
            let name = line
                .trim_start_matches("Class=")
                .split('/')
                .next()
                .unwrap_or("")
                .to_string();
            if !name.is_empty() {
                modules.push((name, MacroModuleType::Class));
            }
        } else if line.starts_with("Document=") {
            let name = line
                .trim_start_matches("Document=")
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
        return Some(data.to_vec());
    }

    let mut out = Vec::new();
    let mut pos = 1usize;
    while pos + 2 <= data.len() {
        let header = u16::from_le_bytes([data[pos], data[pos + 1]]);
        pos += 2;
        let chunk_size = ((header & 0x0FFF) as usize) + 3;
        let compressed = (header & 0x8000) != 0;
        if pos + chunk_size > data.len() {
            break;
        }
        if !compressed {
            out.extend_from_slice(&data[pos..pos + chunk_size]);
            pos += chunk_size;
            continue;
        }

        let chunk_end = pos + chunk_size;
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
                        break;
                    }
                    let token = u16::from_le_bytes([data[pos], data[pos + 1]]);
                    pos += 2;
                    let (offset, length) = decode_copy_token(token, chunk_out.len());
                    for _ in 0..length {
                        if offset == 0 || offset > chunk_out.len() {
                            break;
                        }
                        let b = chunk_out[chunk_out.len() - offset];
                        chunk_out.push(b);
                    }
                }
            }
        }
        out.extend_from_slice(&chunk_out);
    }

    Some(out)
}

fn decode_copy_token(token: u16, decompressed_len: usize) -> (usize, usize) {
    let mut bit_count = 0usize;
    let mut val = if decompressed_len == 0 {
        1
    } else {
        decompressed_len
    };
    while val > 0 {
        bit_count += 1;
        val >>= 1;
    }
    let offset_bits = if bit_count < 4 { 4 } else { bit_count };
    let length_bits = 16 - offset_bits;
    let offset_mask = (1u16 << offset_bits) - 1;
    let length_mask = (1u16 << length_bits) - 1;
    let offset = ((token >> length_bits) & offset_mask) as usize + 1;
    let length = (token & length_mask) as usize + 3;
    (offset, length)
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
    fn vba_decompress_handles_plain_payload_and_invalid_header() {
        assert_eq!(vba_decompress(&[]), None);

        let plain = b"not-compressed";
        assert_eq!(vba_decompress(plain), Some(plain.to_vec()));
    }

    #[test]
    fn vba_decompress_handles_literal_and_copy_tokens() {
        // 0x01 signature + one compressed chunk:
        // flags=0b00000100 => literal 'A', literal 'B', then copy token(offset=2, len=3)
        let encoded = [0x01, 0x02, 0x80, 0x04, b'A', b'B', 0x00, 0x10];
        let out = vba_decompress(&encoded).expect("decompress should succeed");
        assert_eq!(out, b"ABABA");
    }

    #[test]
    fn vba_decompress_gracefully_handles_truncated_chunk() {
        // Declares an uncompressed chunk larger than remaining bytes; parser should stop cleanly.
        let encoded = [0x01, 0x10, 0x00, b'A', b'B'];
        let out = vba_decompress(&encoded).expect("decompress should return partial output");
        assert!(out.is_empty());
    }
}
