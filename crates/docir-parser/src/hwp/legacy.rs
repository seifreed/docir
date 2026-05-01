use super::io::prepare_hwp_stream_data;
use crate::diagnostics::push_warning;
use crate::error::ParseError;
use crate::ole::Cfb;
use docir_core::ir::Diagnostics;
use docir_core::ir::{IRNode, Paragraph, Run};
use docir_core::security::{
    ExternalRefType, ExternalReference, MacroModule, MacroModuleType, MacroProject,
};
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use flate2::read::{DeflateDecoder, ZlibDecoder};
use std::io::Read;

const HWPTAG_BEGIN: u16 = 0x010;
const HWPTAG_DOCUMENT_PROPERTIES: u16 = HWPTAG_BEGIN;
const HWPTAG_PARA_HEADER: u16 = HWPTAG_BEGIN + 50;
const HWPTAG_PARA_TEXT: u16 = HWPTAG_BEGIN + 51;

pub(super) struct HwpHeader {
    pub(super) version: u32,
    pub(super) flags: u32,
}

pub(super) fn parse_file_header(data: &[u8]) -> Result<HwpHeader, ParseError> {
    if data.len() < 40 {
        return Err(ParseError::InvalidStructure(
            "HWP FileHeader too short".to_string(),
        ));
    }
    let signature = &data[..32];
    let signature = String::from_utf8_lossy(signature)
        .trim_matches('\0')
        .to_string();
    if !signature.contains("HWP Document File") {
        return Err(ParseError::InvalidStructure(format!(
            "Invalid HWP signature: {}",
            signature
        )));
    }
    let version = read_u32_le(data, 32)
        .ok_or_else(|| ParseError::InvalidStructure("Missing HWP version".to_string()))?;
    let flags = read_u32_le(data, 36)
        .ok_or_else(|| ParseError::InvalidStructure("Missing HWP flags".to_string()))?;
    Ok(HwpHeader { version, flags })
}

struct HwpRecord<'a> {
    tag_id: u16,
    data: &'a [u8],
}

fn for_each_record<F: FnMut(HwpRecord)>(data: &[u8], mut f: F) -> Result<(), ParseError> {
    let mut offset = 0usize;
    while offset + 4 <= data.len() {
        let header = read_u32_le(data, offset)
            .ok_or_else(|| ParseError::InvalidStructure("Invalid record header".to_string()))?;
        offset += 4;

        let tag_id = (header & 0x3FF) as u16;
        let _level = ((header >> 10) & 0x3FF) as u16;
        let mut size = ((header >> 20) & 0xFFF) as u32;
        if size == 0xFFF {
            if offset + 4 > data.len() {
                return Err(ParseError::InvalidStructure(
                    "Extended record size missing".to_string(),
                ));
            }
            size = read_u32_le(data, offset).ok_or_else(|| {
                ParseError::InvalidStructure("Invalid extended record size".to_string())
            })?;
            offset += 4;
        }

        let end = offset
            .checked_add(size as usize)
            .ok_or_else(|| ParseError::InvalidStructure("Record size overflow".to_string()))?;
        if end > data.len() {
            return Err(ParseError::InvalidStructure(
                "Record size exceeds stream length".to_string(),
            ));
        }
        let payload = &data[offset..end];
        offset = end;
        f(HwpRecord {
            tag_id,
            data: payload,
        });
    }
    Ok(())
}

pub(super) fn parse_docinfo_section_count(data: &[u8]) -> Result<Option<u16>, ParseError> {
    let mut section_count = None;
    for_each_record(data, |rec| {
        if rec.tag_id == HWPTAG_DOCUMENT_PROPERTIES && rec.data.len() >= 2 {
            section_count = read_u16_le(rec.data, 0);
        }
    })?;
    Ok(section_count)
}

pub(super) fn parse_hwp_section_stream(
    data: &[u8],
    source: &str,
    store: &mut IrStore,
) -> Result<Vec<NodeId>, ParseError> {
    let mut paragraphs: Vec<NodeId> = Vec::new();
    let mut current_para: Option<Paragraph> = None;

    for_each_record(data, |rec| match rec.tag_id {
        HWPTAG_PARA_HEADER => {
            if let Some(para) = current_para.take() {
                let para_id = para.id;
                store.insert(IRNode::Paragraph(para));
                paragraphs.push(para_id);
            }
            let mut para = Paragraph::new();
            para.span = Some(SourceSpan::new(source));
            let _text_count = parse_para_text_count(rec.data);
            current_para = Some(para);
        }
        HWPTAG_PARA_TEXT => {
            let text = decode_hwp_text(rec.data);
            if !text.is_empty() {
                let mut run = Run::new(text);
                run.span = Some(SourceSpan::new(source));
                let run_id = run.id;
                store.insert(IRNode::Run(run));
                if let Some(para) = current_para.as_mut() {
                    para.runs.push(run_id);
                } else {
                    let mut para = Paragraph::new();
                    para.span = Some(SourceSpan::new(source));
                    para.runs.push(run_id);
                    current_para = Some(para);
                }
            }
        }
        _ => {}
    })?;

    if let Some(para) = current_para.take() {
        let para_id = para.id;
        store.insert(IRNode::Paragraph(para));
        paragraphs.push(para_id);
    }

    Ok(paragraphs)
}

fn parse_para_text_count(data: &[u8]) -> Option<u32> {
    let mut count = read_u32_le(data, 0)?;
    if count & 0x8000_0000 != 0 {
        count &= 0x7FFF_FFFF;
    }
    Some(count)
}

fn decode_hwp_text(data: &[u8]) -> String {
    if data.is_empty() {
        return String::new();
    }
    let text = if data.len().is_multiple_of(2) {
        decode_utf16le(data)
    } else {
        String::from_utf8_lossy(data).to_string()
    };
    sanitize_text(&text)
}

fn sanitize_text(value: &str) -> String {
    value
        .chars()
        .map(|c| if c <= '\u{001F}' { ' ' } else { c })
        .collect::<String>()
        .trim()
        .to_string()
}

fn decode_utf16le(data: &[u8]) -> String {
    let mut units = Vec::with_capacity(data.len() / 2);
    for chunk in data.chunks(2) {
        if chunk.len() == 2 {
            units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
    }
    String::from_utf16_lossy(&units)
}

pub(super) fn maybe_decompress_stream(
    data: &[u8],
    compressed: bool,
    source: &str,
) -> Result<Vec<u8>, ParseError> {
    if !compressed {
        return Ok(data.to_vec());
    }
    if data.len() < 2 {
        return Ok(data.to_vec());
    }
    let mut out = Vec::new();
    let zlib_result = {
        let mut decoder = ZlibDecoder::new(data);
        decoder.read_to_end(&mut out)
    };
    if zlib_result.is_ok() {
        return Ok(out);
    }
    out.clear();
    let mut decoder = DeflateDecoder::new(data);
    decoder.read_to_end(&mut out).map_err(|e| {
        ParseError::InvalidStructure(format!("Failed to decompress HWP stream {}: {e}", source))
    })?;
    Ok(out)
}

pub(super) fn parse_default_jscript(
    data: &[u8],
    store: &mut IrStore,
    source: &str,
) -> Option<NodeId> {
    let (name, source_code) = parse_jscript_stream(data)?;
    let mut module = MacroModule::new(name, MacroModuleType::Standard);
    module.source_code = Some(source_code);
    module.span = Some(SourceSpan::new(source));
    let module_id = module.id;
    store.insert(IRNode::MacroModule(module));

    let mut project = MacroProject::new();
    project.name = Some("HWP Script".to_string());
    project.modules.push(module_id);
    project.span = Some(SourceSpan::new(source));
    let project_id = project.id;
    store.insert(IRNode::MacroProject(project));

    Some(project_id)
}

fn parse_jscript_stream(data: &[u8]) -> Option<(String, String)> {
    if data.len() < 8 {
        return None;
    }
    let mut offset = 4;
    let mut strings = Vec::new();
    for _ in 0..3 {
        let (value, used) = read_len_string(data, offset)?;
        offset += used;
        if !value.is_empty() {
            strings.push(value);
        }
    }
    let name = strings
        .first()
        .cloned()
        .unwrap_or_else(|| "DefaultJScript".to_string());
    let source = strings.last().cloned().unwrap_or_else(String::new);
    if source.is_empty() {
        None
    } else {
        Some((name, source))
    }
}

const MAX_URL_LENGTH: usize = 4096;

fn extract_urls(text: &str) -> Vec<String> {
    let mut out = Vec::new();
    let patterns: [&[u8]; 5] = [b"http://", b"https://", b"file://", b"ftp://", b"mailto:"];
    let bytes = text.as_bytes();
    let mut idx = 0usize;
    while idx < bytes.len() {
        let mut matched = false;
        for pat in &patterns {
            if bytes[idx..].starts_with(pat) {
                matched = true;
                break;
            }
        }
        if matched {
            let mut end = idx;
            while end < bytes.len() && end - idx < MAX_URL_LENGTH {
                let ch = bytes[end];
                if ch.is_ascii_whitespace() || ch == b'"' || ch == b'\'' || ch == b')' || ch == b'>'
                {
                    break;
                }
                end += 1;
            }
            if end > idx {
                let url = String::from_utf8_lossy(&bytes[idx..end]).to_string();
                out.push(url);
            }
            idx = end;
        } else {
            idx += 1;
        }
    }
    out
}

pub(super) fn scan_hwp_external_refs(
    cfb: &Cfb,
    stream_names: &[String],
    header_ctx: &super::builder::HwpHeaderContext<'_>,
    diagnostics: &mut Diagnostics,
) -> Vec<ExternalReference> {
    let mut refs = Vec::new();
    for path in stream_names {
        if !is_hwp_section_stream(path) {
            continue;
        }
        collect_external_refs_from_stream(cfb, path, header_ctx, diagnostics, &mut refs);
    }
    refs
}

fn is_hwp_section_stream(path: &str) -> bool {
    path.starts_with("BodyText/Section")
}

fn collect_external_refs_from_stream(
    cfb: &Cfb,
    path: &str,
    header_ctx: &super::builder::HwpHeaderContext<'_>,
    diagnostics: &mut Diagnostics,
    refs: &mut Vec<ExternalReference>,
) {
    let Some(data) = cfb.read_stream(path) else {
        return;
    };
    let Some(prepared) = prepare_hwp_stream_data(
        &data,
        header_ctx.encrypted,
        header_ctx.hwp_password,
        header_ctx.force_parse,
        header_ctx.try_raw_encrypted,
        path,
        diagnostics,
    ) else {
        return;
    };

    let decompressed = match maybe_decompress_stream(&prepared, header_ctx.compressed, path) {
        Ok(bytes) => bytes,
        Err(err) => {
            push_warning(
                diagnostics,
                "HWP_DECOMPRESS_FAIL",
                err.to_string(),
                Some(path),
            );
            return;
        }
    };

    push_external_refs_from_text(path, &decode_hwp_text(&decompressed), refs);
}

fn push_external_refs_from_text(path: &str, text: &str, refs: &mut Vec<ExternalReference>) {
    for url in extract_urls(text) {
        let mut ext = ExternalReference::new(ExternalRefType::Hyperlink, url);
        ext.span = Some(SourceSpan::new(path));
        refs.push(ext);
    }
}

fn read_len_string(data: &[u8], offset: usize) -> Option<(String, usize)> {
    let len = read_u32_le(data, offset)? as usize;
    let start = offset + 4;
    let mut bytes_len = len;
    if start + bytes_len > data.len() {
        let alt = len * 2;
        if start + alt <= data.len() {
            bytes_len = alt;
        } else {
            return None;
        }
    }
    let bytes = &data[start..start + bytes_len];
    let text = decode_string_bytes(bytes);
    Some((text, 4 + bytes_len))
}

fn decode_string_bytes(bytes: &[u8]) -> String {
    let zero_bytes = bytes.iter().filter(|b| **b == 0).count();
    if bytes.len().is_multiple_of(2) && zero_bytes > 0 {
        decode_utf16le(bytes)
    } else {
        String::from_utf8_lossy(bytes).to_string()
    }
}

fn read_u16_le(data: &[u8], offset: usize) -> Option<u16> {
    if offset + 2 > data.len() {
        return None;
    }
    Some(u16::from_le_bytes([data[offset], data[offset + 1]]))
}

fn read_u32_le(data: &[u8], offset: usize) -> Option<u32> {
    if offset + 4 > data.len() {
        return None;
    }
    Some(u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ]))
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::IRNode;

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
        let paragraphs = parse_hwp_section_stream(&stream, "BodyText/Section0", &mut store)
            .expect("section parse");
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

        let short = maybe_decompress_stream(&[1], true, "short").expect("short");
        assert_eq!(short, vec![1]);

        let err = maybe_decompress_stream(b"not-compressed", true, "bad")
            .expect_err("invalid compressed");
        assert!(err
            .to_string()
            .contains("Failed to decompress HWP stream bad"));
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
        assert!(module
            .source_code
            .as_deref()
            .unwrap_or_default()
            .contains("function"));

        let utf16 = utf16le("AB");
        let mut len_prefixed = Vec::new();
        len_prefixed.extend_from_slice(&4u32.to_le_bytes());
        len_prefixed.extend_from_slice(&utf16);
        let (value, used) = read_len_string(&len_prefixed, 0).expect("len string");
        assert_eq!(value, "AB");
        assert_eq!(used, 8);
    }

    #[test]
    fn extract_urls_finds_multiple_schemes() {
        let text =
            "before https://a.test/x?q=1 and file://tmp/a.bin and mailto:foo@example.test end";
        let urls = extract_urls(text);
        assert_eq!(urls.len(), 3);
        assert_eq!(urls[0], "https://a.test/x?q=1");
        assert_eq!(urls[1], "file://tmp/a.bin");
        assert_eq!(urls[2], "mailto:foo@example.test");
    }
}
