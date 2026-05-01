use super::{FieldInstruction, FieldKind};

pub(super) fn parse_field_instruction(text: &str) -> Option<FieldInstruction> {
    let tokens = tokenize_field_instruction(text);
    if tokens.is_empty() {
        return None;
    }
    let first = tokens[0].to_ascii_uppercase();
    let kind = match first.as_str() {
        "HYPERLINK" => FieldKind::Hyperlink,
        "INCLUDETEXT" => FieldKind::IncludeText,
        "MERGEFIELD" => FieldKind::MergeField,
        "DATE" => FieldKind::Date,
        "REF" => FieldKind::Ref,
        "PAGEREF" => FieldKind::PageRef,
        _ => FieldKind::Unknown,
    };
    let mut args = Vec::new();
    let mut switches = Vec::new();
    for token in tokens.iter().skip(1) {
        if token.starts_with('\\') {
            switches.push(token.trim_start_matches('\\').to_string());
        } else {
            args.push(token.to_string());
        }
    }
    Some(FieldInstruction {
        kind,
        args,
        switches,
    })
}

pub(super) fn parse_hyperlink_instruction(
    text: &str,
) -> Option<(String, Vec<String>, Vec<String>)> {
    let tokens = tokenize_field_instruction(text);
    if tokens.is_empty() || !tokens[0].eq_ignore_ascii_case("HYPERLINK") {
        return None;
    }
    let mut target = None;
    let mut args = Vec::new();
    let mut switches = Vec::new();
    for token in tokens.into_iter().skip(1) {
        if token.starts_with('\\') {
            switches.push(token.trim_start_matches('\\').to_string());
        } else if target.is_none() {
            target = Some(token);
        } else {
            args.push(token);
        }
    }
    target.map(|t| (t, args, switches))
}

pub(super) fn tokenize_field_instruction(text: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    let mut current = String::new();
    let mut in_quotes = false;
    let mut chars = text.chars().peekable();
    while let Some(ch) = chars.next() {
        match ch {
            '"' => {
                in_quotes = !in_quotes;
            }
            '\\' => {
                if !current.is_empty() {
                    tokens.push(current.clone());
                    current.clear();
                }
                let mut switch = String::from("\\");
                while let Some(&c) = chars.peek() {
                    if c.is_whitespace() {
                        break;
                    }
                    if c == '"' {
                        break;
                    }
                    switch.push(c);
                    chars.next();
                }
                tokens.push(switch);
            }
            c if c.is_whitespace() && !in_quotes => {
                if !current.is_empty() {
                    tokens.push(current.clone());
                    current.clear();
                }
            }
            _ => current.push(ch),
        }
    }
    if !current.is_empty() {
        tokens.push(current);
    }
    tokens
}
