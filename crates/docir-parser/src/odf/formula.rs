pub(crate) fn eval_simple_formula(formula: &str) -> Option<f64> {
    if formula.is_empty() {
        return None;
    }
    if formula
        .chars()
        .any(|c| c.is_alphabetic() || c == '[' || c == ':' || c == ';')
    {
        return None;
    }
    let mut chars = formula.chars().peekable();
    parse_expr(&mut chars)
}

pub(crate) fn parse_expr(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> Option<f64> {
    let mut value = parse_term(chars)?;
    loop {
        skip_ws(chars);
        match chars.peek().copied() {
            Some('+') => {
                chars.next();
                value += parse_term(chars)?;
            }
            Some('-') => {
                chars.next();
                value -= parse_term(chars)?;
            }
            _ => break,
        }
    }
    Some(value)
}

pub(crate) fn parse_term(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> Option<f64> {
    let mut value = parse_factor(chars)?;
    loop {
        skip_ws(chars);
        match chars.peek().copied() {
            Some('*') => {
                chars.next();
                value *= parse_factor(chars)?;
            }
            Some('/') => {
                chars.next();
                let denom = parse_factor(chars)?;
                if denom == 0.0 {
                    return None;
                }
                value /= denom;
            }
            _ => break,
        }
    }
    Some(value)
}

pub(crate) fn parse_factor(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> Option<f64> {
    skip_ws(chars);
    if chars.peek() == Some(&'(') {
        chars.next();
        let value = parse_expr(chars)?;
        skip_ws(chars);
        if chars.peek() == Some(&')') {
            chars.next();
            return Some(value);
        }
        return None;
    }
    parse_number(chars)
}

pub(crate) fn parse_number(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> Option<f64> {
    skip_ws(chars);
    let mut s = String::new();
    while let Some(&c) = chars.peek() {
        if c.is_ascii_digit() || c == '.' {
            s.push(c);
            chars.next();
        } else if c == '-' && s.is_empty() {
            s.push(c);
            chars.next();
        } else {
            break;
        }
    }
    if s.is_empty() {
        None
    } else {
        s.parse::<f64>().ok()
    }
}

pub(crate) fn skip_ws(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) {
    while matches!(chars.peek(), Some(' ' | '\t' | '\n' | '\r')) {
        chars.next();
    }
}
