pub(crate) fn normalize_docx_target(target: &str) -> String {
    let mut t = target;
    while t.starts_with("../") {
        t = &t[3..];
    }
    if t.starts_with("./") {
        t = &t[2..];
    }
    if t.starts_with("word/") {
        t.to_string()
    } else {
        format!("word/{}", t.trim_start_matches('/'))
    }
}
