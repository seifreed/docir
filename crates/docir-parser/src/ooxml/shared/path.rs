pub(crate) fn normalize_docx_target(target: &str) -> String {
    let parts: Vec<&str> = target.split('/').collect();
    let mut resolved: Vec<&str> = Vec::new();
    for part in parts {
        match part {
            "." | "" => {}
            ".." => {
                resolved.pop();
            }
            s => resolved.push(s),
        }
    }
    let t = resolved.join("/");
    if t.starts_with("word/") {
        t
    } else {
        format!("word/{}", t.trim_start_matches('/'))
    }
}
