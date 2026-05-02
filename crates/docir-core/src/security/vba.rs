use super::{SuspiciousCall, SuspiciousCallCategory};

/// Known dangerous XLM functions.
pub const DANGEROUS_XLM_FUNCTIONS: &[&str] = &[
    "EXEC",
    "CALL",
    "REGISTER",
    "RUN",
    "FOPEN",
    "FWRITE",
    "FWRITELN",
    "FREAD",
    "FREADLN",
    "FCLOSE",
    "URLDOWNLOADTOFILE",
    "ALERT",
    "HALT",
    "FORMULA",
    "FORMULA.FILL",
    "SET.VALUE",
    "SET.NAME",
];

/// VBA auto-execution procedures.
pub const AUTO_EXEC_PROCEDURES: &[&str] = &[
    // Word
    "AutoOpen",
    "AutoClose",
    "AutoNew",
    "AutoExec",
    "AutoExit",
    "Document_Open",
    "Document_Close",
    "Document_New",
    // Excel
    "Auto_Open",
    "Auto_Close",
    "Workbook_Open",
    "Workbook_BeforeClose",
    "Workbook_Activate",
];

/// Known suspicious VBA API calls.
pub const SUSPICIOUS_VBA_CALLS: &[(&str, SuspiciousCallCategory)] = &[
    ("Shell", SuspiciousCallCategory::ShellExecution),
    ("WScript.Shell", SuspiciousCallCategory::ShellExecution),
    ("ShellExecute", SuspiciousCallCategory::ShellExecution),
    ("CreateObject", SuspiciousCallCategory::ProcessManipulation),
    ("GetObject", SuspiciousCallCategory::ProcessManipulation),
    ("FileSystemObject", SuspiciousCallCategory::FileSystem),
    (
        "Scripting.FileSystemObject",
        SuspiciousCallCategory::FileSystem,
    ),
    ("XMLHTTP", SuspiciousCallCategory::Network),
    ("WinHTTP", SuspiciousCallCategory::Network),
    ("MSXML2", SuspiciousCallCategory::Network),
    ("InternetExplorer", SuspiciousCallCategory::Network),
    ("PowerShell", SuspiciousCallCategory::PowerShell),
    ("Wscript", SuspiciousCallCategory::ShellExecution),
    ("RegRead", SuspiciousCallCategory::Registry),
    ("RegWrite", SuspiciousCallCategory::Registry),
    ("RegDelete", SuspiciousCallCategory::Registry),
    ("Declare Function", SuspiciousCallCategory::WindowsApi),
    ("Declare Sub", SuspiciousCallCategory::WindowsApi),
    ("CallByName", SuspiciousCallCategory::ProcessManipulation),
    ("Chr", SuspiciousCallCategory::Obfuscation),
    ("ChrW", SuspiciousCallCategory::Obfuscation),
    ("Base64", SuspiciousCallCategory::Obfuscation),
    ("StrReverse", SuspiciousCallCategory::Obfuscation),
    ("Environ", SuspiciousCallCategory::ShellExecution),
];

#[derive(Debug, Clone, Default)]
pub struct VbaAnalysis {
    pub procedures: Vec<String>,
    pub suspicious_calls: Vec<SuspiciousCall>,
    pub auto_exec_procedures: Vec<String>,
}

/// Public API entrypoint: is_dangerous_xlm_function.
pub fn is_dangerous_xlm_function(name: &str) -> bool {
    let upper = name.to_ascii_uppercase();
    DANGEROUS_XLM_FUNCTIONS.iter().any(|&item| item == upper)
}

/// Public API entrypoint: analyze_vba_source.
pub fn analyze_vba_source(source: &str) -> VbaAnalysis {
    let mut analysis = VbaAnalysis::default();

    for (line_num, line) in source.lines().enumerate() {
        let raw = line.trim();
        if raw.is_empty() {
            continue;
        }

        let lower = raw.to_ascii_lowercase();
        if let Some(procedure) = parse_vba_procedure_name(raw) {
            analysis.procedures.push(procedure);
        }

        let is_comment =
            raw.starts_with('\'') || raw.starts_with("Rem ") || raw.starts_with("Rem:");
        if is_comment {
            continue;
        }

        for (pattern, category) in SUSPICIOUS_VBA_CALLS {
            let pattern_lower = pattern.to_ascii_lowercase();
            for pos in lower.match_indices(&pattern_lower) {
                let (idx, _) = pos;
                let before_ok = idx == 0
                    || !lower
                        .as_bytes()
                        .get(idx - 1)
                        .is_some_and(|b| b.is_ascii_alphanumeric() || *b == b'_');
                let after_ok = idx + pattern_lower.len() >= lower.len()
                    || !lower
                        .as_bytes()
                        .get(idx + pattern_lower.len())
                        .is_some_and(|b| b.is_ascii_alphanumeric() || *b == b'_');
                if before_ok && after_ok {
                    analysis.suspicious_calls.push(SuspiciousCall {
                        name: (*pattern).to_string(),
                        category: *category,
                        line: Some(line_num as u32 + 1),
                    });
                    break;
                }
            }
        }

        if let Some(proc_name) = parse_vba_procedure_name(raw) {
            for proc in AUTO_EXEC_PROCEDURES {
                if proc_name.eq_ignore_ascii_case(proc) {
                    analysis.auto_exec_procedures.push(proc.to_string());
                }
            }
        }
    }

    analysis
}

fn parse_vba_procedure_name(raw: &str) -> Option<String> {
    let mut tokens: Vec<&str> = raw.split_whitespace().collect();
    strip_visibility_modifier(&mut tokens);
    if tokens.len() < 2 {
        return None;
    }

    let keyword = tokens[0];
    if !(keyword.eq_ignore_ascii_case("sub") || keyword.eq_ignore_ascii_case("function")) {
        return None;
    }

    let name = tokens[1].split('(').next().unwrap_or("");
    if name.is_empty() {
        None
    } else {
        Some(name.to_string())
    }
}

fn strip_visibility_modifier(tokens: &mut Vec<&str>) {
    if tokens.len() >= 2
        && (tokens[0].eq_ignore_ascii_case("private")
            || tokens[0].eq_ignore_ascii_case("public")
            || tokens[0].eq_ignore_ascii_case("friend")
            || tokens[0].eq_ignore_ascii_case("static"))
    {
        tokens.remove(0);
    }
}

#[cfg(test)]
mod tests {
    use super::{analyze_vba_source, is_dangerous_xlm_function, AUTO_EXEC_PROCEDURES};
    use crate::security::SuspiciousCallCategory;

    #[test]
    fn dangerous_xlm_function_check_is_case_insensitive() {
        assert!(is_dangerous_xlm_function("exec"));
        assert!(is_dangerous_xlm_function("Call"));
        assert!(!is_dangerous_xlm_function("SUM"));
    }

    #[test]
    fn analyze_vba_source_extracts_procedures_and_suspicious_calls() {
        let source = r#"
            Sub AutoOpen()
                Shell "calc.exe"
                Dim x: x = Chr(65)
            End Sub

            Private Function BuildValue()
                Set o = CreateObject("WScript.Shell")
                BuildValue = "ok"
            End Function
        "#;

        let analysis = analyze_vba_source(source);
        assert!(analysis.procedures.iter().any(|p| p == "AutoOpen"));
        assert!(analysis.procedures.iter().any(|p| p == "BuildValue"));
        assert!(analysis
            .auto_exec_procedures
            .iter()
            .any(|p| p == "AutoOpen"));
        assert!(analysis.suspicious_calls.iter().any(|c| {
            c.name == "Shell" && c.category == SuspiciousCallCategory::ShellExecution
        }));
        assert!(analysis.suspicious_calls.iter().any(|c| {
            c.name == "CreateObject" && c.category == SuspiciousCallCategory::ProcessManipulation
        }));
        assert!(analysis
            .suspicious_calls
            .iter()
            .any(|c| { c.name == "Chr" && c.category == SuspiciousCallCategory::Obfuscation }));
        assert!(analysis.suspicious_calls.iter().all(|c| c.line.is_some()));
    }

    #[test]
    fn analyze_vba_source_handles_non_matching_lines() {
        let source = r#"
            Option Explicit
            Dim value As String
            value = "hello"
        "#;
        let analysis = analyze_vba_source(source);
        assert!(analysis.procedures.is_empty());
        assert!(analysis.suspicious_calls.is_empty());
        assert!(analysis.auto_exec_procedures.is_empty());
    }

    #[test]
    fn auto_exec_procedures_constant_has_expected_entries() {
        assert!(AUTO_EXEC_PROCEDURES.contains(&"AutoOpen"));
        assert!(AUTO_EXEC_PROCEDURES.contains(&"Workbook_Open"));
    }
}
