use super::*;
use crate::test_support::build_test_cfb;
use docir_core::ir::IRNode;
use flate2::Compression;
use flate2::write::ZlibEncoder;
use std::io::Write;

fn patch_cfb_fat_entry(bytes: &[u8], fat_index: u32, value: u32) -> Vec<u8> {
    let mut out = bytes.to_vec();
    let sector_size = 1usize << u16::from_le_bytes([out[0x1E], out[0x1F]]);
    let first_fat_sector = u32::from_le_bytes([out[0x4C], out[0x4D], out[0x4E], out[0x4F]]);
    let fat_offset = sector_size + first_fat_sector as usize * sector_size + fat_index as usize * 4;
    out[fat_offset..fat_offset + 4].copy_from_slice(&value.to_le_bytes());
    out
}

fn make_record(tag_id: u16, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    if payload.len() < 0x0FFF {
        let header = (tag_id as u32) | ((payload.len() as u32) << 20);
        out.extend_from_slice(&header.to_le_bytes());
    } else {
        let header = (tag_id as u32) | (0x0FFFu32 << 20);
        out.extend_from_slice(&header.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    }
    out.extend_from_slice(payload);
    out
}

fn utf16le(text: &str) -> Vec<u8> {
    text.encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .collect::<Vec<_>>()
}

#[test]
fn parse_file_header_validates_signature_and_fields() {
    let mut data = vec![0u8; 40];
    let sig = b"HWP Document File";
    data[..sig.len()].copy_from_slice(sig);
    data[32..36].copy_from_slice(&5u32.to_le_bytes());
    data[36..40].copy_from_slice(&1u32.to_le_bytes());

    let parsed = parse_file_header(&data).expect("valid header");
    assert_eq!(parsed.version, 5);
    assert_eq!(parsed.flags, 1);

    assert!(parse_file_header(&data[..20]).is_err());

    let mut invalid = data.clone();
    invalid[..7].copy_from_slice(b"invalid");
    assert!(parse_file_header(&invalid).is_err());
}

#[test]
fn parse_docinfo_section_count_supports_normal_and_extended_records() {
    let section_payload = 3u16.to_le_bytes();
    let rec = make_record(HWPTAG_DOCUMENT_PROPERTIES, &section_payload);
    let count = parse_docinfo_section_count(&rec)
        .expect("docinfo parse")
        .expect("count");
    assert_eq!(count, 3);

    let large_payload = vec![0x41u8; 0x1005];
    let rec = make_record(HWPTAG_DOCUMENT_PROPERTIES + 1, &large_payload);
    assert_eq!(parse_docinfo_section_count(&rec).unwrap(), None);
}

#[test]
fn parse_hwp_section_stream_emits_paragraphs_and_runs() {
    let mut stream = Vec::new();
    stream.extend(make_record(HWPTAG_PARA_HEADER, &4u32.to_le_bytes()));
    stream.extend(make_record(HWPTAG_PARA_TEXT, &utf16le("Hello")));
    stream.extend(make_record(HWPTAG_PARA_HEADER, &2u32.to_le_bytes()));
    stream.extend(make_record(HWPTAG_PARA_TEXT, b"World"));

    let mut store = IrStore::new();
    let paragraphs =
        parse_hwp_section_stream(&stream, "BodyText/Section0", &mut store).expect("section parse");
    assert_eq!(paragraphs.len(), 2);

    let first_para = match store.get(paragraphs[0]).expect("paragraph node") {
        IRNode::Paragraph(p) => p,
        _ => panic!("expected paragraph"),
    };
    assert_eq!(first_para.runs.len(), 1);
    let first_run = match store.get(first_para.runs[0]).expect("run node") {
        IRNode::Run(r) => r,
        _ => panic!("expected run"),
    };
    assert_eq!(first_run.text, "Hello");
}

#[test]
fn decoding_and_sanitizing_helpers_cover_utf16_and_controls() {
    assert_eq!(decode_hwp_text(&[]), "");
    assert_eq!(decode_hwp_text(b"abc"), "abc");
    assert_eq!(decode_hwp_text(&utf16le("A\u{0001}B")), "A B");
    assert_eq!(sanitize_text("  x\n"), "x");
    assert_eq!(decode_string_bytes(&utf16le("Name")), "Name");
    assert_eq!(decode_string_bytes(b"plain"), "plain");
}

#[test]
fn decompression_handles_passthrough_and_invalid_payloads() {
    let passthrough = maybe_decompress_stream(b"abc", false, "s").expect("passthrough");
    assert_eq!(passthrough, b"abc");

    let err = maybe_decompress_stream(&[1], true, "short").expect_err("short compressed stream");
    assert!(
        err.to_string()
            .contains("Compressed HWP stream short is too short")
    );

    let err =
        maybe_decompress_stream(b"not-compressed", true, "bad").expect_err("invalid compressed");
    assert!(
        err.to_string()
            .contains("Failed to decompress HWP stream bad")
    );
}

#[test]
fn decompression_rejects_output_above_resource_limit() {
    let mut encoder = ZlibEncoder::new(Vec::new(), Compression::fast());
    encoder
        .write_all(&vec![b'A'; MAX_HWP_DECOMPRESSED_SIZE + 1])
        .expect("write compressed payload");
    let compressed = encoder.finish().expect("finish compressed payload");

    let err = maybe_decompress_stream(&compressed, true, "oversized")
        .expect_err("decompressed output above limit must fail");
    assert!(matches!(err, ParseError::ResourceLimit(_)));
}

#[test]
fn default_jscript_and_len_strings_are_parsed() {
    let mut data = vec![0, 0, 0, 0];
    let name = b"Module1";
    let source = b"function x(){return 1;}";
    data.extend_from_slice(&(name.len() as u32).to_le_bytes());
    data.extend_from_slice(name);
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&(source.len() as u32).to_le_bytes());
    data.extend_from_slice(source);

    let mut store = IrStore::new();
    let project_id =
        parse_default_jscript(&data, &mut store, "Scripts/DefaultJScript").expect("project id");
    let project = match store.get(project_id).expect("project node") {
        IRNode::MacroProject(p) => p,
        _ => panic!("expected macro project"),
    };
    assert_eq!(project.name.as_deref(), Some("HWP Script"));
    assert_eq!(project.modules.len(), 1);
    let module = match store.get(project.modules[0]).expect("module node") {
        IRNode::MacroModule(m) => m,
        _ => panic!("expected macro module"),
    };
    assert_eq!(module.name, "Module1");
    assert!(
        module
            .source_code
            .as_deref()
            .unwrap_or_default()
            .contains("function")
    );

    let utf16 = utf16le("AB");
    let mut len_prefixed = Vec::new();
    len_prefixed.extend_from_slice(&4u32.to_le_bytes());
    len_prefixed.extend_from_slice(&utf16);
    let (value, used) = read_len_string(&len_prefixed, 0).expect("len string");
    assert_eq!(value, "AB");
    assert_eq!(used, 8);
}

#[test]
fn default_jscript_rejects_malformed_len_strings() {
    let mut store = IrStore::new();
    let err = parse_default_jscript(b"\0\0\0\0\x20\0", &mut store, "Scripts/DefaultJScript")
        .expect_err("malformed script stream should fail");

    assert!(matches!(err, ParseError::InvalidStructure(_)));
}

#[test]
fn extract_urls_finds_multiple_schemes() {
    let text = "before https://a.test/x?q=1 and file://tmp/a.bin and mailto:foo@example.test end";
    let urls = extract_urls(text);
    assert_eq!(urls.len(), 3);
    assert_eq!(urls[0], "https://a.test/x?q=1");
    assert_eq!(urls[1], "file://tmp/a.bin");
    assert_eq!(urls[2], "mailto:foo@example.test");
}

#[test]
fn external_ref_scan_reports_corrupt_section_stream_read() {
    let section = vec![b'A'; 5000];
    let base = build_test_cfb(&[("Section0", &section)]);
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
    let mut refs = Vec::new();

    collect_external_refs_from_stream(&cfb, "Section0", &header_ctx, &mut diagnostics, &mut refs);

    assert!(refs.is_empty());
    assert!(diagnostics.entries.iter().any(|e| {
        e.code == "HWP_STREAM_READ_FAIL" && e.message.contains("OLE sector out of bounds")
    }));
}
