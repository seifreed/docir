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
