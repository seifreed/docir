//! Security indicator detection utilities.

use docir_core::security::{
    SuspiciousCall, SuspiciousCallCategory, ThreatIndicator, ThreatIndicatorType, ThreatLevel,
    SUSPICIOUS_VBA_CALLS,
};
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
    let source_upper = source.to_uppercase();

    for (pattern, category) in SUSPICIOUS_VBA_CALLS {
        let pattern_upper = pattern.to_uppercase();
        if source_upper.contains(&pattern_upper) {
            // Find line number
            let line = source
                .lines()
                .enumerate()
                .find(|(_, line)| line.to_uppercase().contains(&pattern_upper))
                .map(|(i, _)| i as u32 + 1);

            calls.push(SuspiciousCall {
                name: pattern.to_string(),
                category: *category,
                line,
            });
        }
    }

    calls
}

/// Checks if a URL is potentially suspicious.
pub fn is_suspicious_url(url: &str) -> bool {
    let url_lower = url.to_lowercase();

    // Check for suspicious TLDs
    let suspicious_tlds = [".ru", ".cn", ".tk", ".ml", ".ga", ".cf"];
    for tld in &suspicious_tlds {
        if url_lower.ends_with(tld) || url_lower.contains(&format!("{}/", tld)) {
            return true;
        }
    }

    // Check for IP addresses (potential C2)
    let parts: Vec<&str> = url_lower
        .trim_start_matches("http://")
        .trim_start_matches("https://")
        .split('/')
        .next()
        .unwrap_or("")
        .split('.')
        .collect();

    if parts.len() == 4 && parts.iter().all(|p| p.parse::<u8>().is_ok()) {
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
        if url_lower.contains(pattern) {
            return true;
        }
    }

    false
}

/// Checks if a procedure name is an auto-execute trigger.
pub fn is_auto_exec_procedure(name: &str) -> bool {
    use docir_core::security::AUTO_EXEC_PROCEDURES;

    let name_lower = name.to_lowercase();
    AUTO_EXEC_PROCEDURES
        .iter()
        .any(|p| p.to_lowercase() == name_lower)
}

/// Checks if a formula contains dangerous XLM functions.
pub fn contains_dangerous_xlm(formula: &str) -> Vec<String> {
    use docir_core::security::DANGEROUS_XLM_FUNCTIONS;

    let formula_upper = formula.to_uppercase();
    DANGEROUS_XLM_FUNCTIONS
        .iter()
        .filter(|f| formula_upper.contains(&f.to_uppercase()))
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
        assert!(!contains_dangerous_xlm("=SUM(A1:A10)").is_empty() == false);
        assert!(!contains_dangerous_xlm("=EXEC(\"cmd /c calc\")").is_empty());
        assert!(!contains_dangerous_xlm("=CALL(\"kernel32\",\"WinExec\")").is_empty());
    }
}
