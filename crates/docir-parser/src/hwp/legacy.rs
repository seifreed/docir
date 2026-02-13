use super::*;
use flate2::read::{DeflateDecoder, ZlibDecoder};

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
    level: u16,
    size: u32,
    data: &'a [u8],
}

fn for_each_record<F: FnMut(HwpRecord)>(data: &[u8], mut f: F) -> Result<(), ParseError> {
    let mut offset = 0usize;
    while offset + 4 <= data.len() {
        let header = read_u32_le(data, offset)
            .ok_or_else(|| ParseError::InvalidStructure("Invalid record header".to_string()))?;
        offset += 4;

        let tag_id = (header & 0x3FF) as u16;
        let level = ((header >> 10) & 0x3FF) as u16;
        let mut size = ((header >> 20) & 0xFFF) as u32;
        if size == 0xFFF {
            size = read_u32_le(data, offset).ok_or_else(|| {
                ParseError::InvalidStructure("Invalid extended record size".to_string())
            })?;
            offset += 4;
        }

        let end = offset + size as usize;
        if end > data.len() {
            return Err(ParseError::InvalidStructure(
                "Record size exceeds stream length".to_string(),
            ));
        }
        let payload = &data[offset..end];
        offset = end;
        f(HwpRecord {
            tag_id,
            level,
            size,
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
    let text = if data.len() % 2 == 0 {
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
        .get(0)
        .cloned()
        .unwrap_or_else(|| "DefaultJScript".to_string());
    let source = strings.last().cloned().unwrap_or_else(String::new);
    if source.is_empty() {
        None
    } else {
        Some((name, source))
    }
}

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
            while end < bytes.len() {
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
    compressed: bool,
    encrypted: bool,
    password: Option<&str>,
    force_parse: bool,
    try_raw_encrypted: bool,
    diagnostics: &mut Diagnostics,
) -> Vec<ExternalReference> {
    let mut refs = Vec::new();
    for path in stream_names {
        if !path.starts_with("BodyText/Section") {
            continue;
        }
        if let Some(data) = cfb.read_stream(path) {
            let data = match prepare_hwp_stream_data(
                &data,
                encrypted,
                password,
                force_parse,
                try_raw_encrypted,
                path,
                diagnostics,
            ) {
                Some(bytes) => bytes,
                None => continue,
            };
            let data = match maybe_decompress_stream(&data, compressed, path) {
                Ok(bytes) => bytes,
                Err(err) => {
                    push_warning(
                        diagnostics,
                        "HWP_DECOMPRESS_FAIL",
                        err.to_string(),
                        Some(path),
                    );
                    continue;
                }
            };
            let text = decode_hwp_text(&data);
            for url in extract_urls(&text) {
                let mut ext = ExternalReference::new(ExternalRefType::Hyperlink, url);
                ext.span = Some(SourceSpan::new(path));
                refs.push(ext);
            }
        }
    }
    refs
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
    if bytes.len() % 2 == 0 && zero_bytes > 0 {
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
