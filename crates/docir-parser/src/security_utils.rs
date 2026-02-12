use docir_core::security::{DdeField, DdeFieldType};
use docir_core::types::SourceSpan;

pub(crate) fn parse_dde_formula(
    formula: &str,
    location: SourceSpan,
    require_prefix: bool,
) -> Option<DdeField> {
    let trimmed = formula.trim();
    let upper = trimmed.to_ascii_uppercase();

    let (field_type, args_start) = if require_prefix {
        if upper.starts_with("DDEAUTO(") {
            (DdeFieldType::DdeAuto, "DDEAUTO(".len())
        } else if upper.starts_with("DDE(") {
            (DdeFieldType::Dde, "DDE(".len())
        } else {
            return None;
        }
    } else if let Some(idx) = upper.find("DDEAUTO(") {
        (DdeFieldType::DdeAuto, idx + "DDEAUTO(".len())
    } else if let Some(idx) = upper.find("DDE(") {
        (DdeFieldType::Dde, idx + "DDE(".len())
    } else {
        return None;
    };

    let args_end = trimmed.rfind(')')?;
    if args_end <= args_start {
        return None;
    }
    let args = &trimmed[args_start..args_end];
    let parts = split_formula_args(args);
    if parts.is_empty() {
        return None;
    }

    let application = normalize_arg(parts.get(0)?);
    let topic = parts
        .get(1)
        .map(|v| normalize_arg(v))
        .filter(|v| !v.is_empty());
    let item = parts
        .get(2)
        .map(|v| normalize_arg(v))
        .filter(|v| !v.is_empty());

    Some(DdeField {
        field_type,
        application,
        topic,
        item,
        instruction: formula.to_string(),
        location: Some(location),
    })
}

fn split_formula_args(args: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut current = String::new();
    let mut in_quotes = false;
    let mut quote_char = '\0';
    for ch in args.chars() {
        if (ch == '"' || ch == '\'') && (!in_quotes || ch == quote_char) {
            in_quotes = !in_quotes;
            if in_quotes {
                quote_char = ch;
            }
            continue;
        }
        if !in_quotes && (ch == ';' || ch == ',') {
            let trimmed = current.trim();
            if !trimmed.is_empty() {
                out.push(trimmed.to_string());
            }
            current.clear();
            continue;
        }
        current.push(ch);
    }
    let trimmed = current.trim();
    if !trimmed.is_empty() {
        out.push(trimmed.to_string());
    }
    out
}

fn normalize_arg(value: &str) -> String {
    value
        .trim()
        .trim_matches('"')
        .trim_matches('\'')
        .to_string()
}
