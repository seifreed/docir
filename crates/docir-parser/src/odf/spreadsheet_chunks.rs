use super::{attr_value, Event, Reader};

#[derive(Clone)]
pub(super) struct OdfTableChunk {
    pub(super) start: usize,
    pub(super) end: usize,
    pub(super) bytes: Vec<u8>,
}

pub(super) fn extract_spreadsheet_table_chunks(xml: &[u8]) -> Vec<OdfTableChunk> {
    let Some((start, end)) = find_spreadsheet_range(xml) else {
        return Vec::new();
    };
    let mut chunks = Vec::new();
    let mut pos = start;
    while let Some(idx) = find_subslice(xml, b"<table:table", pos, end) {
        if let Some(next) = xml.get(idx + b"<table:table".len()) {
            if *next == b'-' {
                pos = idx + 1;
                continue;
            }
        }
        let Some(tag_end) = find_tag_end(xml, idx + 1, end) else {
            break;
        };
        let self_closing = is_self_closing_tag(xml, idx, tag_end);
        let chunk_end = if self_closing {
            tag_end
        } else {
            let Some(close_start) = find_subslice(xml, b"</table:table>", tag_end + 1, end) else {
                break;
            };
            close_start + b"</table:table>".len() - 1
        };
        if chunk_end >= idx {
            let bytes = xml[idx..=chunk_end].to_vec();
            chunks.push(OdfTableChunk {
                start: idx,
                end: chunk_end,
                bytes,
            });
        }
        pos = chunk_end.saturating_add(1);
    }
    chunks
}

pub(super) fn table_name_from_chunk(chunk: &[u8], sheet_id: u32) -> String {
    let mut reader = Reader::from_reader(std::io::Cursor::new(chunk));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"table:table" => {
                return attr_value(&e, b"table:name")
                    .unwrap_or_else(|| format!("Sheet{}", sheet_id));
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
    format!("Sheet{}", sheet_id)
}

fn find_spreadsheet_range(xml: &[u8]) -> Option<(usize, usize)> {
    let start = find_subslice(xml, b"<office:spreadsheet", 0, xml.len())?;
    let tag_end = find_tag_end(xml, start + 1, xml.len())?;
    let end_tag = find_subslice(xml, b"</office:spreadsheet>", tag_end + 1, xml.len())?;
    let end = end_tag + b"</office:spreadsheet>".len();
    Some((tag_end + 1, end))
}

fn find_subslice(haystack: &[u8], needle: &[u8], start: usize, end: usize) -> Option<usize> {
    let mut i = start;
    let limit = end.saturating_sub(needle.len());
    while i <= limit {
        if &haystack[i..i + needle.len()] == needle {
            return Some(i);
        }
        i += 1;
    }
    None
}

fn find_tag_end(xml: &[u8], start: usize, end: usize) -> Option<usize> {
    let mut i = start;
    let mut quote: Option<u8> = None;
    while i < end {
        let b = xml[i];
        if let Some(q) = quote {
            if b == q {
                quote = None;
            }
        } else if b == b'"' || b == b'\'' {
            quote = Some(b);
        } else if b == b'>' {
            return Some(i);
        }
        i += 1;
    }
    None
}

fn is_self_closing_tag(xml: &[u8], start: usize, end: usize) -> bool {
    let mut i = end.saturating_sub(1);
    while i > start {
        let b = xml[i];
        if b == b'/' {
            return true;
        }
        if !b.is_ascii_whitespace() {
            break;
        }
        i = i.saturating_sub(1);
    }
    false
}
