//! Security indicator detection utilities.

use crate::vba::{
    contains_dangerous_xlm as detect_dangerous_xlm, is_auto_exec_procedure as is_auto_exec_name,
    scan_vba_source as scan_vba_calls,
};
use docir_core::security::{SuspiciousCall, ThreatIndicator, ThreatIndicatorType, ThreatLevel};
use docir_core::types::NodeId;
use std::net::IpAddr;

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
/// Uses word-boundary matching to avoid false positives from substring matches
/// (e.g., "architecture" matching "Chr").
pub fn scan_vba_source(source: &str) -> Vec<SuspiciousCall> {
    scan_vba_calls(source)
}

/// Checks if a URL is potentially suspicious.
pub fn is_suspicious_url(url: &str) -> bool {
    let url_lower = url.to_lowercase();

    // Decode percent-encoded sequences to detect evasion techniques.
    let decoded = percent_decode_lower(&url_lower);

    // Check for suspicious TLDs against the host portion of the decoded URL.
    let host = extract_host(&decoded);
    let suspicious_tlds = [".ru", ".cn", ".tk", ".ml", ".ga", ".cf"];
    for tld in &suspicious_tlds {
        if host.ends_with(tld) {
            return true;
        }
    }

    // Check for IP addresses (potential C2) using the decoded host portion.
    if is_ip_address(&host) {
        return true;
    }

    // Check for UNC paths (\\server\share) commonly used in LNK attacks.
    if decoded.starts_with("\\\\") || decoded.starts_with("//") {
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
        if host_matches_domain(&host, pattern) {
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

    // Strip port, path, query, and fragment.
    let host_port = no_auth
        .split(['/', '\\', '?', '#'])
        .next()
        .unwrap_or(no_auth)
        .trim_end_matches('.');
    let host = if let Some(bracketed) = host_port.strip_prefix('[') {
        bracketed.split(']').next().unwrap_or(bracketed)
    } else {
        host_port.split(':').next().unwrap_or(host_port)
    }
    .trim_end_matches('.');

    host.to_string()
}

fn host_matches_domain(host: &str, domain: &str) -> bool {
    host == domain
        || host
            .strip_suffix(domain)
            .is_some_and(|prefix| prefix.ends_with('.'))
}

/// Checks if a string is an IPv4 address.
fn is_ip_address(host: &str) -> bool {
    let parts: Vec<&str> = host.split('.').collect();
    if parts.len() == 4 && parts.iter().all(|p| p.parse::<u8>().is_ok()) {
        return true;
    }
    if host.parse::<IpAddr>().is_ok() {
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
    is_auto_exec_name(name)
}

/// Checks if a formula contains dangerous XLM functions.
/// Uses word-boundary matching to avoid false positives from substring matches.
pub fn contains_dangerous_xlm(formula: &str) -> Vec<String> {
    detect_dangerous_xlm(formula)
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
        assert!(is_suspicious_url("http://[2001:db8::1]/payload.exe"));
        assert!(is_suspicious_url("http://evil.ru/malware.doc"));
        assert!(is_suspicious_url("http://evil.ru?download=1"));
        assert!(is_suspicious_url("http://192.168.1.1#payload"));
        assert!(is_suspicious_url(r"http://evil.ru\malware.doc"));
        assert!(is_suspicious_url("http://evil.ru%5cmalware.doc"));
        assert!(is_suspicious_url(r"http://192.168.1.1\payload.exe"));
        assert!(is_suspicious_url("https://pastebin.com/raw/abc123"));
        assert!(is_suspicious_url("https://pastebin.com?raw=abc123"));
        assert!(is_suspicious_url("https://pastebin.com%3fraw=abc123"));
        assert!(is_suspicious_url("https://pastebin.com./raw/abc123"));
        assert!(is_suspicious_url("http://evil.ru./malware.doc"));
        assert!(!is_suspicious_url("https://www.microsoft.com/docs"));
        assert!(!is_suspicious_url("https://notpastebin.com/raw/abc123"));
        assert!(!is_suspicious_url("https://example.com?next=pastebin.com"));
        assert!(!is_suspicious_url(
            "https://example.com/path/pastebin.com/report"
        ));
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
