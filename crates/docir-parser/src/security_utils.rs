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

    let args_end = find_matching_paren(trimmed, args_start.saturating_sub(1))?;
    if args_end <= args_start {
        return None;
    }
    let args = &trimmed[args_start..args_end];
    let parts = split_formula_args(args);
    if parts.is_empty() {
        return None;
    }

    let application = normalize_arg(parts.first()?);
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

fn find_matching_paren(s: &str, open_pos: usize) -> Option<usize> {
    let chars: Vec<char> = s.chars().collect();
    if open_pos >= chars.len() || chars[open_pos] != '(' {
        return None;
    }
    let mut depth = 1usize;
    let mut in_quotes = false;
    let mut quote_char = '\0';
    let mut prev_ch = '\0';
    for (i, &ch) in chars.iter().enumerate().skip(open_pos + 1) {
        if !in_quotes {
            if (ch == '"' || ch == '\'') && (quote_char == '\0' || ch == quote_char) {
                in_quotes = true;
                quote_char = ch;
            } else if ch == '(' {
                depth += 1;
            } else if ch == ')' {
                depth -= 1;
                if depth == 0 {
                    return Some(i);
                }
            }
        } else if ch == quote_char {
            // Doubled quotes ("") inside a quoted string are escape sequences, not closers
            if prev_ch == ch {
                // Already handled: the previous char was the same quote, so this is a doubled quote
                // that we already skipped. Do nothing.
            } else {
                // Look ahead: if next char is same quote, it's a doubled quote (escaped)
                let next_is_same = chars.get(i + 1) == Some(&ch);
                if next_is_same {
                    // This is the opening quote of a doubled pair; stay in quotes
                } else {
                    in_quotes = false;
                    quote_char = '\0';
                }
            }
        }
        prev_ch = ch;
    }
    None
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
