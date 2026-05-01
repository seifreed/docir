use docir_core::ir::{ExtensionPart, ExtensionPartKind, MediaType};

/// Public API entrypoint: classify_media_type.
pub fn classify_media_type(path: &str) -> MediaType {
    let ext = path.rsplit('.').next().unwrap_or("").to_ascii_lowercase();
    match ext.as_str() {
        "jpg" | "jpeg" | "png" | "gif" | "bmp" | "tif" | "tiff" | "webp" => MediaType::Image,
        "mp3" | "wav" | "m4a" | "aac" | "ogg" => MediaType::Audio,
        "mp4" | "mov" | "avi" | "mkv" | "webm" => MediaType::Video,
        _ => MediaType::Other,
    }
}

/// Public API entrypoint: legacy_extension_part.
pub fn legacy_extension_part(path: &str, size_bytes: u64) -> ExtensionPart {
    ExtensionPart::new(path, size_bytes, ExtensionPartKind::Legacy)
}
