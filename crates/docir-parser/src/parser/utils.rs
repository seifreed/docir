use docir_core::types::NodeId;
use std::collections::HashMap;

pub(crate) fn resolve_media_asset(
    media_by_path: &HashMap<String, NodeId>,
    target: &str,
) -> Option<NodeId> {
    if let Some(id) = media_by_path.get(target) {
        return Some(*id);
    }
    let trimmed = target.trim_start_matches('/');
    for (path, id) in media_by_path {
        if path.ends_with(trimmed) || trimmed.ends_with(path) {
            return Some(*id);
        }
    }
    None
}

pub(crate) fn find_stream_case<'a>(streams: &'a [String], name: &str) -> Option<&'a str> {
    let target = name.to_ascii_uppercase();
    streams
        .iter()
        .find(|s| s.to_ascii_uppercase() == target)
        .map(|s| s.as_str())
}
