//! Shared helpers for part registries.

pub(crate) fn matches_pattern(path: &str, pattern: &str) -> bool {
    if !pattern.contains('*') {
        return path == pattern;
    }

    let mut parts = pattern.splitn(2, '*');
    let prefix = parts.next().unwrap_or("");
    let suffix = parts.next().unwrap_or("");

    if !prefix.is_empty() && !path.starts_with(prefix) {
        return false;
    }
    if !suffix.is_empty() && !path.ends_with(suffix) {
        return false;
    }
    true
}
