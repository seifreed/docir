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
        return parse_finite_seconds(stripped);
    }
    if trimmed.starts_with("PT") && trimmed.ends_with('S') {
        let inner = trimmed.strip_prefix("PT")?.strip_suffix('S')?;
        return parse_finite_seconds(inner);
    }
    None
}

fn parse_finite_seconds(value: &str) -> Option<u32> {
    let seconds = value.parse::<f32>().ok()?;
    if seconds.is_finite() {
        Some((seconds * 1000.0).round() as u32)
    } else {
        None
    }
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
