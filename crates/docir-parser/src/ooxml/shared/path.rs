use crate::ooxml::relationships::looks_like_external_target;

pub(crate) fn normalize_docx_target(target: &str) -> String {
    if looks_like_external_target(target) {
        return target.to_string();
    }

    let normalized_target = target.replace('\\', "/");
    let parts: Vec<&str> = normalized_target.split('/').collect();
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

#[cfg(test)]
mod tests {
    use super::normalize_docx_target;

    #[test]
    fn test_normalize_docx_target_handles_backslash_separators() {
        assert_eq!(
            normalize_docx_target(r"..\media\image1.png"),
            "word/media/image1.png"
        );
    }

    #[test]
    fn test_normalize_docx_target_preserves_external_targets() {
        assert_eq!(
            normalize_docx_target("https://example.test/image.png"),
            "https://example.test/image.png"
        );
    }
}
