use docir_core::security::{DdeField, DdeFieldType};

pub(super) fn parse_dde_instruction(instruction: &str) -> Option<DdeField> {
    let field_type = dde_field_type(instruction)?;

    let parts = extract_quoted_parts(instruction);
    let application = parts.first().cloned().unwrap_or_default();
    if application.is_empty() {
        // DDE field without quoted arguments: still flag it with a generic description.
        return Some(DdeField {
            field_type,
            application: "unknown".to_string(),
            topic: None,
            item: None,
            instruction: instruction.to_string(),
            location: None,
        });
    }
    let topic = parts.get(1).cloned().unwrap_or_default();
    let item = parts.get(2).cloned().unwrap_or_default();

    Some(DdeField {
        field_type,
        application,
        topic: if topic.is_empty() { None } else { Some(topic) },
        item: if item.is_empty() { None } else { Some(item) },
        instruction: instruction.to_string(),
        location: None,
    })
}

fn dde_field_type(instruction: &str) -> Option<DdeFieldType> {
    let keyword = instruction.split_whitespace().next()?;
    if keyword.eq_ignore_ascii_case("DDEAUTO") {
        Some(DdeFieldType::DdeAuto)
    } else if keyword.eq_ignore_ascii_case("DDE") {
        Some(DdeFieldType::Dde)
    } else {
        None
    }
}

fn extract_quoted_parts(input: &str) -> Vec<String> {
    let mut parts = Vec::new();
    let mut current = String::new();
    let mut in_quotes = false;
    let mut chars = input.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '"' {
            if in_quotes {
                // Check for escaped quote ""
                if chars.peek() == Some(&'"') {
                    chars.next();
                    current.push('"');
                } else {
                    in_quotes = false;
                    parts.push(current.clone());
                    current.clear();
                }
            } else {
                in_quotes = true;
            }
            continue;
        }
        if in_quotes {
            current.push(ch);
        }
    }
    // Handle unclosed quote: include accumulated text as a part
    if in_quotes && !current.is_empty() {
        parts.push(current);
    }
    parts
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::security::DdeFieldType;

    #[test]
    fn parse_dde_instruction_extracts_parts_and_type() {
        let parsed =
            parse_dde_instruction(r#"DDEAUTO "cmd" "/c calc" "A1""#).expect("expected DDE parse");
        assert_eq!(parsed.field_type, DdeFieldType::DdeAuto);
        assert_eq!(parsed.application, "cmd");
        assert_eq!(parsed.topic.as_deref(), Some("/c calc"));
        assert_eq!(parsed.item.as_deref(), Some("A1"));
    }

    #[test]
    fn parse_dde_instruction_handles_missing_quoted_parts() {
        let parsed =
            parse_dde_instruction(r#"DDEAUTO "winword" "" """#).expect("expected DDE parse");
        assert_eq!(parsed.field_type, DdeFieldType::DdeAuto);
        assert_eq!(parsed.application, "winword");
        assert_eq!(parsed.topic, None);
        assert_eq!(parsed.item, None);
    }

    #[test]
    fn parse_dde_instruction_rejects_empty_application() {
        // DDE with empty quoted app still flags as "unknown" to avoid false negatives
        let result = parse_dde_instruction(r#"DDEAUTO "" "/c calc" "A1""#);
        assert!(result.is_some());
        assert_eq!(result.unwrap().application, "unknown");
        assert!(parse_dde_instruction("DDE").is_some());
        assert!(parse_dde_instruction("HYPERLINK").is_none());
    }

    #[test]
    fn parse_dde_instruction_ignores_non_dde_fields() {
        assert!(parse_dde_instruction(r#"HYPERLINK "https://example.test""#).is_none());
        assert!(parse_dde_instruction(r#"NOTDDE "cmd" "/c calc""#).is_none());
        assert!(parse_dde_instruction(r#"SOMEDDEAUTO "cmd" "/c calc""#).is_none());
    }
}
