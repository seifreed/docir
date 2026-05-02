pub(super) fn classify_payload(
    data: &[u8],
    file_name: Option<&str>,
) -> (&'static str, &'static str) {
    if data.starts_with(b"MZ") {
        return ("pe", "application/vnd.microsoft.portable-executable");
    }
    if data.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]) {
        return ("office", "application/x-cfb");
    }
    if data.starts_with(b"PK\x03\x04") {
        if file_name
            .map(|name| {
                let lower = name.to_ascii_lowercase();
                lower.ends_with(".docx")
                    || lower.ends_with(".docm")
                    || lower.ends_with(".xlsx")
                    || lower.ends_with(".xlsm")
                    || lower.ends_with(".pptx")
                    || lower.ends_with(".pptm")
            })
            .unwrap_or(false)
        {
            return ("office", "application/vnd.openxmlformats-officedocument");
        }
        return ("zip", "application/zip");
    }
    if data.starts_with(b"%PDF-") {
        return ("pdf", "application/pdf");
    }
    if looks_like_text(data) {
        return ("script", "text/plain");
    }
    ("unknown", "application/octet-stream")
}

pub(super) fn classify_media_asset(path: &str, data: &[u8]) -> (&'static str, &'static str) {
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".png") {
        return ("image", "image/png");
    }
    if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        return ("image", "image/jpeg");
    }
    if lower.ends_with(".gif") {
        return ("image", "image/gif");
    }
    if lower.ends_with(".bmp") {
        return ("image", "image/bmp");
    }
    if lower.ends_with(".tif") || lower.ends_with(".tiff") {
        return ("image", "image/tiff");
    }
    if lower.ends_with(".webp") {
        return ("image", "image/webp");
    }
    if lower.ends_with(".wav") {
        return ("audio", "audio/wav");
    }
    if lower.ends_with(".mp3") {
        return ("audio", "audio/mpeg");
    }
    if lower.ends_with(".mp4") {
        return ("video", "video/mp4");
    }
    classify_payload(data, Some(path))
}

fn looks_like_text(data: &[u8]) -> bool {
    if data.is_empty() {
        return false;
    }
    let sample = &data[..data.len().min(256)];
    sample
        .iter()
        .all(|byte| byte.is_ascii_whitespace() || byte.is_ascii_graphic())
}
