//! Field instruction parsing helpers for DOCX.

use docir_core::ir::{FieldInstruction, FieldKind};

pub(super) fn parse_field_instruction(instr: &str) -> Option<FieldInstruction> {
    let decoded = unescape_xml_entities(instr);
    let tokens = tokenize_field_instruction(&decoded);
    if tokens.is_empty() {
        return None;
    }
    let kind = match tokens[0].as_str() {
        "HYPERLINK" => FieldKind::Hyperlink,
        "INCLUDETEXT" => FieldKind::IncludeText,
        "INCLUDEPICTURE" => FieldKind::IncludePicture,
        "MERGEFIELD" => FieldKind::MergeField,
        "DATE" => FieldKind::Date,
        "REF" => FieldKind::Ref,
        "PAGEREF" => FieldKind::PageRef,
        "DDE" => FieldKind::Dde,
        "DDEAUTO" => FieldKind::DdeAuto,
        "AUTOTEXT" => FieldKind::AutoText,
        "AUTOCORRECT" => FieldKind::AutoCorrect,
        _ => FieldKind::Unknown,
    };
    let mut args = Vec::new();
    let mut switches = Vec::new();
    for tok in tokens.into_iter().skip(1) {
        if tok.starts_with('\\') {
            switches.push(normalize_switch(&tok));
        } else {
            args.push(tok);
        }
    }
    for sw in extract_switches(&decoded) {
        let sw = normalize_switch(&sw);
        if !switches.contains(&sw) {
            switches.push(sw);
        }
    }
    if decoded.contains('\t') && !switches.iter().any(|s| s == "\\t") {
        switches.push("\\t".to_string());
    }
    if matches!(kind, FieldKind::Hyperlink) {
        let mut normalized_args = Vec::new();
        for arg in args {
            if arg.len() == 1 && arg.chars().all(|c| c.is_ascii_alphabetic()) {
                let sw = format!("\\{arg}");
                if !switches.contains(&sw) {
                    switches.push(sw);
                }
            } else {
                normalized_args.push(arg);
            }
        }
        return Some(FieldInstruction {
            kind,
            args: normalized_args,
            switches,
        });
    }
    Some(FieldInstruction {
        kind,
        args,
        switches,
    })
}

fn tokenize_field_instruction(instr: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    let mut current = String::new();
    let mut in_quotes = false;
    let mut chars = instr.chars().peekable();
    while let Some(ch) = chars.next() {
        match ch {
            '"' => {
                in_quotes = !in_quotes;
            }
            ' ' | '\t' | '\r' | '\n' if !in_quotes => {
                if !current.is_empty() {
                    tokens.push(current.clone());
                    current.clear();
                }
                while matches!(chars.peek(), Some(' ' | '\t' | '\r' | '\n')) {
                    chars.next();
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

fn extract_switches(instr: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut in_quotes = false;
    let chars: Vec<char> = instr.chars().collect();
    let mut i = 0usize;
    while i < chars.len() {
        let ch = chars[i];
        if ch == '"' {
            in_quotes = !in_quotes;
            i += 1;
            continue;
        }
        if !in_quotes && ch == '\\' {
            let mut j = i + 1;
            if j >= chars.len() || chars[j].is_whitespace() {
                i += 1;
                continue;
            }
            let mut token = String::new();
            token.push('\\');
            while j < chars.len() && !chars[j].is_whitespace() {
                token.push(chars[j]);
                j += 1;
            }
            out.push(token);
            i = j;
            continue;
        }
        i += 1;
    }
    out
}

fn unescape_xml_entities(value: &str) -> String {
    value
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
}

fn normalize_switch(value: &str) -> String {
    let mut out = value.to_string();
    while out.starts_with("\\\\") {
        out.remove(0);
    }
    out
}
