use super::MediaType;

pub(super) fn parse_duration_ms(value: &str) -> Option<u32> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return None;
    }
    if let Some(stripped) = trimmed.strip_suffix("ms") {
        return stripped.parse::<u32>().ok();
    }
    if let Some(stripped) = trimmed.strip_suffix('s') {
        return stripped
            .parse::<f32>()
            .ok()
            .map(|v| (v * 1000.0).round() as u32);
    }
    if trimmed.starts_with("PT") && trimmed.ends_with('S') {
        let inner = trimmed.trim_start_matches("PT").trim_end_matches('S');
        return inner
            .parse::<f32>()
            .ok()
            .map(|v| (v * 1000.0).round() as u32);
    }
    None
}

pub(super) fn classify_media_type(path: &str, media: &str) -> Option<MediaType> {
    let lower_media = media.to_ascii_lowercase();
    if lower_media.starts_with("image/") {
        return Some(MediaType::Image);
    }
    if lower_media.starts_with("audio/") {
        return Some(MediaType::Audio);
    }
    if lower_media.starts_with("video/") {
        return Some(MediaType::Video);
    }
    if lower_media.starts_with("application/") {
        let lower_path = path.to_ascii_lowercase();
        if lower_path.ends_with(".ogg") || lower_path.ends_with(".oga") {
            return Some(MediaType::Audio);
        }
        if lower_path.ends_with(".ogv") {
            return Some(MediaType::Video);
        }
    }
    None
}
