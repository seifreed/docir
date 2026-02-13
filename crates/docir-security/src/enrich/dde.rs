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
            if !in_quotes && !current.is_empty() {
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
