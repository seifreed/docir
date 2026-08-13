use super::MediaType;
use crate::error::ParseError;

pub(super) fn parse_duration_ms(value: &str) -> Result<Option<u32>, ParseError> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return Ok(None);
    }
    if let Some(stripped) = trimmed.strip_suffix("ms") {
        return stripped
            .parse::<u32>()
            .map(Some)
            .map_err(|err| invalid_duration(trimmed, err));
    }
    if let Some(stripped) = trimmed.strip_suffix('s') {
        return parse_finite_seconds(stripped);
    }
    if trimmed.starts_with("PT") && trimmed.ends_with('S') {
        let inner = trimmed
            .strip_prefix("PT")
            .and_then(|value| value.strip_suffix('S'))
            .ok_or_else(|| {
                ParseError::InvalidStructure(format!("Invalid ODF duration '{trimmed}'"))
            })?;
        return parse_finite_seconds(inner);
    }
    Err(ParseError::InvalidStructure(format!(
        "Invalid ODF duration '{trimmed}'"
    )))
}

fn parse_finite_seconds(value: &str) -> Result<Option<u32>, ParseError> {
    let seconds = value
        .parse::<f64>()
        .map_err(|err| invalid_duration(value, err))?;
    if seconds.is_finite() && seconds >= 0.0 {
        let milliseconds = (seconds * 1000.0).round();
        if milliseconds.is_finite() && milliseconds <= u32::MAX as f64 {
            return Ok(Some(milliseconds as u32));
        }
    } else {
        return Err(ParseError::InvalidStructure(format!(
            "Invalid ODF duration '{value}'"
        )));
    }
    Err(ParseError::InvalidStructure(format!(
        "ODF duration exceeds u32 milliseconds: '{value}'"
    )))
}

fn invalid_duration(value: &str, err: impl std::fmt::Display) -> ParseError {
    ParseError::InvalidStructure(format!("Invalid ODF duration '{value}': {err}"))
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

#[cfg(test)]
mod tests {
    use super::parse_duration_ms;
    use crate::error::ParseError;

    #[test]
    fn parse_duration_ms_rejects_overflowing_seconds() {
        let err = parse_duration_ms("4294968s").expect_err("duration must fit u32 ms");

        assert!(
            matches!(err, ParseError::InvalidStructure(message) if message.contains("exceeds"))
        );
    }
}
