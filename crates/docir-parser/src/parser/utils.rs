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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolve_media_asset_matches_direct_and_suffix_paths() {
        let mut media = HashMap::new();
        let id_direct = NodeId::from_raw(11);
        let id_suffix = NodeId::from_raw(12);
        media.insert("word/media/image1.png".to_string(), id_direct);
        media.insert("ppt/media/photo.jpg".to_string(), id_suffix);

        assert_eq!(
            resolve_media_asset(&media, "word/media/image1.png"),
            Some(id_direct)
        );
        assert_eq!(
            resolve_media_asset(&media, "/ppt/media/photo.jpg"),
            Some(id_suffix)
        );
        assert_eq!(
            resolve_media_asset(&media, "media/photo.jpg"),
            Some(id_suffix)
        );
        assert_eq!(resolve_media_asset(&media, "media/missing.bin"), None);
    }

    #[test]
    fn find_stream_case_is_case_insensitive() {
        let streams = vec![
            "WordDocument".to_string(),
            "Data".to_string(),
            "1Table".to_string(),
        ];
        assert_eq!(
            find_stream_case(&streams, "worddocument"),
            Some("WordDocument")
        );
        assert_eq!(find_stream_case(&streams, "1TABLE"), Some("1Table"));
        assert_eq!(find_stream_case(&streams, "Missing"), None);
    }
}
