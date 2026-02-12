use quick_xml::events::BytesStart;

pub(crate) fn attr_value(e: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == name {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

pub(crate) fn attr_value_by_suffix(e: &BytesStart<'_>, suffixes: &[&[u8]]) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        for suffix in suffixes {
            if key.ends_with(suffix) {
                if let Ok(value) = attr.unescape_value() {
                    return Some(value.to_string());
                }
                return Some(String::from_utf8_lossy(&attr.value).to_string());
            }
        }
    }
    None
}

pub(crate) fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(pos) => &name[pos + 1..],
        None => name,
    }
}
