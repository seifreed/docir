use docir_core::security::{DdeField, DdeFieldType};

pub(super) fn parse_dde_instruction(instruction: &str) -> Option<DdeField> {
    let upper = instruction.to_ascii_uppercase();
    if !upper.contains("DDE") {
        return None;
    }
    let field_type = if upper.contains("DDEAUTO") {
        DdeFieldType::DdeAuto
    } else {
        DdeFieldType::Dde
    };

    let parts = extract_quoted_parts(instruction);
    let application = parts.first().cloned().unwrap_or_default();
    if application.is_empty() {
        return None;
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

fn extract_quoted_parts(input: &str) -> Vec<String> {
    let mut parts = Vec::new();
    let mut current = String::new();
    let mut in_quotes = false;
    for ch in input.chars() {
        if ch == '"' {
            in_quotes = !in_quotes;
            if !in_quotes {
                parts.push(current.clone());
                current.clear();
            }
            continue;
        }
        if in_quotes {
            current.push(ch);
        }
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
        assert!(parse_dde_instruction(r#"DDEAUTO "" "/c calc" "A1""#).is_none());
        assert!(parse_dde_instruction("DDE").is_none());
    }

    #[test]
    fn parse_dde_instruction_ignores_non_dde_fields() {
        assert!(parse_dde_instruction(r#"HYPERLINK "https://example.test""#).is_none());
    }
}
