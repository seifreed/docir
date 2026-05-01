//! Security indicator detection utilities.

use crate::policy::{AUTO_EXEC_PROCEDURES, DANGEROUS_XLM_FUNCTIONS, SUSPICIOUS_VBA_CALLS};
use docir_core::security::{SuspiciousCall, ThreatIndicator, ThreatIndicatorType, ThreatLevel};
use docir_core::types::NodeId;

/// Builds a threat indicator with standardized fields.
pub fn make_indicator(
    indicator_type: ThreatIndicatorType,
    severity: ThreatLevel,
    description: impl Into<String>,
    location: Option<String>,
    node_id: Option<NodeId>,
) -> ThreatIndicator {
    ThreatIndicator {
        indicator_type,
        severity,
        description: description.into(),
        location,
        node_id,
    }
}

/// Scans VBA source code for suspicious API calls.
pub fn scan_vba_source(source: &str) -> Vec<SuspiciousCall> {
    let mut calls = Vec::new();

    for (pattern, category) in SUSPICIOUS_VBA_CALLS {
        let pattern_upper = pattern.to_uppercase();
        for (i, line) in source.lines().enumerate() {
            if line.to_uppercase().contains(&pattern_upper) {
                calls.push(SuspiciousCall {
                    name: pattern.to_string(),
                    category: *category,
                    line: Some(i as u32 + 1),
                });
            }
        }
    }

    calls
}

/// Checks if a URL is potentially suspicious.
pub fn is_suspicious_url(url: &str) -> bool {
    let url_lower = url.to_lowercase();

    // Decode percent-encoded sequences to detect evasion techniques.
    let decoded = percent_decode_lower(&url_lower);

    // Check for suspicious TLDs against the decoded URL.
    let suspicious_tlds = [".ru", ".cn", ".tk", ".ml", ".ga", ".cf"];
    for tld in &suspicious_tlds {
        if decoded.ends_with(tld) || decoded.contains(&format!("{}/", tld)) {
            return true;
        }
    }

    // Check for IP addresses (potential C2) using the decoded host portion.
    let host = extract_host(&decoded);
    if is_ip_address(&host) {
        return true;
    }

    // Check for UNC paths (\\server\share) commonly used in LNK attacks.
    if url_lower.starts_with("\\\\") || url_lower.starts_with("//") {
        return true;
    }

    // Check for suspicious patterns
    let suspicious_patterns = [
        "pastebin.com",
        "bit.ly",
        "tinyurl.com",
        "raw.githubusercontent.com",
        "ngrok.io",
        "serveo.net",
    ];

    for pattern in &suspicious_patterns {
        if decoded.contains(pattern) {
            return true;
        }
    }

    false
}

/// Decodes percent-encoded sequences in a lowercase URL for matching.
/// Only decodes the common evasion characters (dots, slashes, colons).
fn percent_decode_lower(input: &str) -> String {
    let mut result = String::with_capacity(input.len());
    let bytes = input.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let Ok(byte_val) = u8::from_str_radix(&input[i + 1..i + 3], 16) {
                result.push(byte_val.to_ascii_lowercase() as char);
                i += 3;
                continue;
            }
        }
        result.push(bytes[i] as char);
        i += 1;
    }
    result
}

/// Extracts the host portion from a URL, stripping scheme, auth, port, and path.
fn extract_host(url: &str) -> String {
    let no_scheme = url
        .strip_prefix("http://")
        .or_else(|| url.strip_prefix("https://"))
        .or_else(|| url.strip_prefix("ftp://"))
        .or_else(|| url.strip_prefix("smb://"))
        .or_else(|| url.strip_prefix("sftp://"))
        .or_else(|| url.strip_prefix("file://"))
        .or_else(|| url.strip_prefix("ldap://"))
        .or_else(|| url.strip_prefix("ldaps://"))
        .unwrap_or(url);

    // Strip userinfo (user:pass@)
    let no_auth = no_scheme.split('@').next_back().unwrap_or(no_scheme);

    // Strip port and path
    let host_port = no_auth.split('/').next().unwrap_or(no_auth);
    let host = host_port.split(':').next().unwrap_or(host_port);

    host.to_string()
}

/// Checks if a string is an IPv4 address.
fn is_ip_address(host: &str) -> bool {
    let parts: Vec<&str> = host.split('.').collect();
    if parts.len() == 4 && parts.iter().all(|p| p.parse::<u8>().is_ok()) {
        return true;
    }
    // Hex IP addresses (0xC0A80101 style)
    let hex_prefix = host.strip_prefix("0x").or_else(|| host.strip_prefix("0X"));
    if let Some(hex_part) = hex_prefix {
        if u32::from_str_radix(hex_part, 16).is_ok() {
            return true;
        }
    }
    false
}

/// Checks if a procedure name is an auto-execute trigger.
pub fn is_auto_exec_procedure(name: &str) -> bool {
    AUTO_EXEC_PROCEDURES
        .iter()
        .any(|p| p.eq_ignore_ascii_case(name))
}

/// Checks if a formula contains dangerous XLM functions.
/// Uses word-boundary matching to avoid false positives from substring matches.
pub fn contains_dangerous_xlm(formula: &str) -> Vec<String> {
    let formula_upper = formula.to_uppercase();
    DANGEROUS_XLM_FUNCTIONS
        .iter()
        .filter(|f| {
            let name = f.to_uppercase();
            // Check that the function name appears at a word boundary:
            // either at the start, after =, after (, after ,, or after a space.
            for pos in formula_upper.match_indices(&name) {
                let (idx, _) = pos;
                let before_ok = idx == 0
                    || matches!(
                        formula_upper.as_bytes().get(idx - 1),
                        Some(b'=' | b'(' | b',' | b' ' | b'\t')
                    );
                let after_ok = idx + name.len() >= formula_upper.len()
                    || matches!(
                        formula_upper.as_bytes().get(idx + name.len()),
                        Some(b'(' | b',' | b')' | b' ' | b'\t' | b'.' | b'!')
                    );
                if before_ok && after_ok {
                    return true;
                }
            }
            false
        })
        .map(|f| f.to_string())
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_scan_vba_source() {
        let source = r#"
Sub AutoOpen()
    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    ws.Run "cmd /c calc.exe"
End Sub
"#;
        let calls = scan_vba_source(source);
        assert!(calls.iter().any(|c| c.name.contains("CreateObject")));
        assert!(calls
            .iter()
            .any(|c| c.name.contains("WScript.Shell") || c.name.contains("Wscript")));
    }

    #[test]
    fn test_is_suspicious_url() {
        assert!(is_suspicious_url("http://192.168.1.1/payload.exe"));
        assert!(is_suspicious_url("http://evil.ru/malware.doc"));
        assert!(is_suspicious_url("https://pastebin.com/raw/abc123"));
        assert!(!is_suspicious_url("https://www.microsoft.com/docs"));
    }

    #[test]
    fn test_is_auto_exec_procedure() {
        assert!(is_auto_exec_procedure("AutoOpen"));
        assert!(is_auto_exec_procedure("Document_Open"));
        assert!(is_auto_exec_procedure("Workbook_Open"));
        assert!(!is_auto_exec_procedure("MyCustomSub"));
    }

    #[test]
    fn test_contains_dangerous_xlm() {
        assert!(contains_dangerous_xlm("=SUM(A1:A10)").is_empty());
        assert!(!contains_dangerous_xlm("=EXEC(\"cmd /c calc\")").is_empty());
        assert!(!contains_dangerous_xlm("=CALL(\"kernel32\",\"WinExec\")").is_empty());
    }
}
