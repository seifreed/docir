use docir_core::ir::TextAlignment;

pub(crate) fn parse_text_alignment(value: &str) -> Option<TextAlignment> {
    let value = value.trim().to_ascii_lowercase();
    match value.as_str() {
        "left" => Some(TextAlignment::Left),
        "center" => Some(TextAlignment::Center),
        "right" => Some(TextAlignment::Right),
        "justify" | "justified" => Some(TextAlignment::Justify),
        "distribute" => Some(TextAlignment::Distribute),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_text_alignment_supports_known_variants() {
        assert_eq!(parse_text_alignment("left"), Some(TextAlignment::Left));
        assert_eq!(
            parse_text_alignment(" center "),
            Some(TextAlignment::Center)
        );
        assert_eq!(parse_text_alignment("RIGHT"), Some(TextAlignment::Right));
        assert_eq!(
            parse_text_alignment("justify"),
            Some(TextAlignment::Justify)
        );
        assert_eq!(
            parse_text_alignment("justified"),
            Some(TextAlignment::Justify)
        );
        assert_eq!(
            parse_text_alignment("distribute"),
            Some(TextAlignment::Distribute)
        );
        assert_eq!(parse_text_alignment("unknown"), None);
    }
}
