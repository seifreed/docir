use super::{Event, Reader};
use crate::error::ParseError;
use crate::xml_utils::{XmlScanControl, local_name, scan_xml_events, try_attr_value_by_suffix};

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
    while let Some((idx, tag_end, prefix)) = find_start_tag_by_local(xml, b"table", pos, end) {
        let self_closing = is_self_closing_tag(xml, idx, tag_end);
        let chunk_end = if self_closing {
            tag_end
        } else {
            let Some(close_start) = find_matching_end_tag(xml, prefix, b"table", tag_end + 1, end)
            else {
                break;
            };
            let Some(close_end) = find_tag_end(xml, close_start + 2, end) else {
                break;
            };
            close_end
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

pub(super) fn table_name_from_chunk(chunk: &[u8], sheet_id: u32) -> Result<String, ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(chunk));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut table_name = None;
    scan_xml_events(&mut reader, &mut buf, "content.xml", |event| match event {
        Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"table" => {
            table_name = Some(
                try_attr_value_by_suffix(&e, &[b":name"], "content.xml")?
                    .unwrap_or_else(|| format!("Sheet{}", sheet_id)),
            );
            Ok(XmlScanControl::Break)
        }
        _ => Ok(XmlScanControl::Continue),
    })?;

    Ok(table_name.unwrap_or_else(|| format!("Sheet{}", sheet_id)))
}

fn find_spreadsheet_range(xml: &[u8]) -> Option<(usize, usize)> {
    let (_, tag_end, prefix) = find_start_tag_by_local(xml, b"spreadsheet", 0, xml.len())?;
    let end_tag = find_end_tag(xml, prefix, b"spreadsheet", tag_end + 1, xml.len())?;
    let end = find_tag_end(xml, end_tag + 2, xml.len())? + 1;
    Some((tag_end + 1, end))
}

fn find_start_tag_by_local<'a>(
    xml: &'a [u8],
    local: &[u8],
    start: usize,
    end: usize,
) -> Option<(usize, usize, &'a [u8])> {
    let mut i = start;
    while i < end {
        if xml[i] != b'<' || matches!(xml.get(i + 1), Some(b'/') | Some(b'!') | Some(b'?')) {
            i += 1;
            continue;
        }
        let name_start = i + 1;
        let name_end = find_name_end(xml, name_start, end)?;
        let (prefix, name_local) = split_qualified_name(&xml[name_start..name_end]);
        if name_local == local {
            let tag_end = find_tag_end(xml, name_end, end)?;
            return Some((i, tag_end, prefix));
        }
        i += 1;
    }
    None
}

fn find_end_tag(
    xml: &[u8],
    prefix: &[u8],
    local: &[u8],
    start: usize,
    end: usize,
) -> Option<usize> {
    let mut i = start;
    while i < end {
        if xml.get(i..i + 2) != Some(b"</") {
            i += 1;
            continue;
        }
        let name_start = i + 2;
        let name_end = find_name_end(xml, name_start, end)?;
        let (candidate_prefix, candidate_local) = split_qualified_name(&xml[name_start..name_end]);
        if candidate_prefix == prefix && candidate_local == local {
            return Some(i);
        }
        i += 1;
    }
    None
}

fn find_matching_end_tag(
    xml: &[u8],
    prefix: &[u8],
    local: &[u8],
    start: usize,
    end: usize,
) -> Option<usize> {
    let mut depth = 1usize;
    let mut i = start;
    while i < end {
        if xml[i] != b'<' {
            i += 1;
            continue;
        }
        match xml.get(i + 1) {
            Some(b'/') => {
                let name_start = i + 2;
                let name_end = find_name_end(xml, name_start, end)?;
                let (candidate_prefix, candidate_local) =
                    split_qualified_name(&xml[name_start..name_end]);
                if candidate_prefix == prefix && candidate_local == local {
                    depth = depth.saturating_sub(1);
                    if depth == 0 {
                        return Some(i);
                    }
                }
            }
            Some(b'!') | Some(b'?') => {}
            Some(_) => {
                let name_start = i + 1;
                let name_end = find_name_end(xml, name_start, end)?;
                let tag_end = find_tag_end(xml, name_end, end)?;
                let (candidate_prefix, candidate_local) =
                    split_qualified_name(&xml[name_start..name_end]);
                if candidate_prefix == prefix
                    && candidate_local == local
                    && !is_self_closing_tag(xml, i, tag_end)
                {
                    depth += 1;
                }
                i = tag_end;
            }
            None => return None,
        }
        i += 1;
    }
    None
}

fn find_name_end(xml: &[u8], start: usize, end: usize) -> Option<usize> {
    let mut i = start;
    while i < end {
        if matches!(xml[i], b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>') {
            return Some(i);
        }
        i += 1;
    }
    None
}

fn split_qualified_name(name: &[u8]) -> (&[u8], &[u8]) {
    if let Some(idx) = name.iter().position(|b| *b == b':') {
        (&name[..idx], &name[idx + 1..])
    } else {
        (b"".as_slice(), name)
    }
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

#[cfg(test)]
mod tests {
    use super::{extract_spreadsheet_table_chunks, table_name_from_chunk};
    use crate::error::ParseError;

    #[test]
    fn extract_spreadsheet_table_chunks_accepts_alternate_prefixes() {
        let xml =
            br#"<pkg:document-content xmlns:pkg="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <pkg:body>
    <pkg:spreadsheet>
      <t:table t:name="Alt1"><t:table-row/></t:table>
      <t:table t:name="Alt2"/>
    </pkg:spreadsheet>
  </pkg:body>
</pkg:document-content>"#;

        let chunks = extract_spreadsheet_table_chunks(xml);

        assert_eq!(chunks.len(), 2);
        assert_eq!(
            table_name_from_chunk(&chunks[0].bytes, 1).expect("table name"),
            "Alt1"
        );
        assert_eq!(
            table_name_from_chunk(&chunks[1].bytes, 2).expect("table name"),
            "Alt2"
        );
    }

    #[test]
    fn extract_spreadsheet_table_chunks_keeps_nested_tables_in_parent_chunk() {
        let xml = br#"<office:document-content>
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Outer">
        <table:table-row>
          <table:table-cell>
            <table:table table:name="Inner"><table:table-row/></table:table>
          </table:table-cell>
        </table:table-row>
        <table:table-row table:style-name="after-inner"/>
      </table:table>
      <table:table table:name="Next"/>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;

        let chunks = extract_spreadsheet_table_chunks(xml);

        assert_eq!(chunks.len(), 2);
        assert_eq!(
            table_name_from_chunk(&chunks[0].bytes, 1).expect("table name"),
            "Outer"
        );
        assert!(
            chunks[0]
                .bytes
                .windows(b"after-inner".len())
                .any(|w| w == b"after-inner")
        );
        assert_eq!(
            table_name_from_chunk(&chunks[1].bytes, 2).expect("table name"),
            "Next"
        );
    }

    #[test]
    fn table_name_from_chunk_reports_invalid_attribute_entity() {
        let chunk = br#"<table:table table:name="Broken &"/>"#;

        let err = table_name_from_chunk(chunk, 1).expect_err("invalid table name entity must fail");
        assert!(matches!(err, ParseError::Xml { file, .. } if file == "content.xml"));
    }
}
