pub(super) fn preferred_output_name(
    suggested_name: Option<&str>,
    index: usize,
    kind: &str,
    mime_type: Option<&str>,
) -> String {
    if let Some(name) = suggested_name.and_then(empty_to_none_str) {
        return sanitize_name(name);
    }
    let ext = match kind {
        "pe" => "exe",
        "pdf" => "pdf",
        "zip" => "zip",
        "office" => "bin",
        "script" => "txt",
        _ => mime_type
            .and_then(|value| value.rsplit('/').next())
            .unwrap_or("bin"),
    };
    format!("artifact_{}.{}", index, ext)
}

pub(super) fn sanitize_name(input: &str) -> String {
    let mut out = String::with_capacity(input.len());
    for ch in input.chars() {
        if ch.is_ascii_alphanumeric() || ch == '.' || ch == '-' || ch == '_' {
            out.push(ch);
        } else {
            out.push('_');
        }
    }
    if out.is_empty() {
        "artifact.bin".to_string()
    } else {
        out
    }
}

pub(super) fn file_name_from_path(path: &str) -> String {
    path.rsplit('/').next().unwrap_or(path).to_string()
}

pub(super) fn sha256_hex(data: &[u8]) -> String {
    docir_security::sha256_hex(data)
}

pub(super) fn assign_sha256(target: &mut Option<String>, data: &[u8], compute_hashes: bool) {
    if compute_hashes {
        *target = Some(sha256_hex(data));
    }
}

pub(super) fn empty_to_none(value: String) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

pub(super) fn empty_to_none_str(value: &str) -> Option<&str> {
    if value.trim().is_empty() {
        None
    } else {
        Some(value)
    }
}
