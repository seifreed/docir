//! VBA and XLM security signature helpers.

use docir_core::security::{SuspiciousCall, SuspiciousCallCategory};

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
    "AutoOpen",
    "AutoClose",
    "AutoNew",
    "AutoExec",
    "AutoExit",
    "Document_Open",
    "Document_Close",
    "Document_New",
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

/// Summary of VBA source analysis.
#[derive(Debug, Clone, Default)]
pub struct VbaAnalysis {
    /// Procedure names declared in the module.
    pub procedures: Vec<String>,
    /// Suspicious calls found in executable lines.
    pub suspicious_calls: Vec<SuspiciousCall>,
    /// Auto-execute procedures declared by the module.
    pub auto_exec_procedures: Vec<String>,
}

/// Checks whether a formula function name is in the dangerous XLM set.
pub fn is_dangerous_xlm_function(name: &str) -> bool {
    let upper = name.to_ascii_uppercase();
    DANGEROUS_XLM_FUNCTIONS.iter().any(|&item| item == upper)
}

/// Scans VBA source code for procedures, auto-exec triggers, and suspicious calls.
pub fn analyze_vba_source(source: &str) -> VbaAnalysis {
    let mut analysis = VbaAnalysis::default();

    for (line_index, line) in source.lines().enumerate() {
        let raw = line.trim();
        if raw.is_empty() {
            continue;
        }
        record_procedure(raw, &mut analysis);
        if is_comment_line(raw) {
            continue;
        }
        analysis
            .suspicious_calls
            .extend(suspicious_calls_in_line(raw, line_index as u32 + 1));
    }

    analysis
}

/// Returns suspicious VBA calls found in source text.
pub fn scan_vba_source(source: &str) -> Vec<SuspiciousCall> {
    analyze_vba_source(source).suspicious_calls
}

/// Checks whether a procedure name is an auto-execute trigger.
pub fn is_auto_exec_procedure(name: &str) -> bool {
    AUTO_EXEC_PROCEDURES
        .iter()
        .any(|p| p.eq_ignore_ascii_case(name))
}

/// Finds dangerous XLM functions in a formula.
pub fn contains_dangerous_xlm(formula: &str) -> Vec<String> {
    let formula_upper = formula.to_uppercase();
    DANGEROUS_XLM_FUNCTIONS
        .iter()
        .filter(|function| formula_contains_xlm_function(&formula_upper, function))
        .map(|function| function.to_string())
        .collect()
}

fn record_procedure(raw: &str, analysis: &mut VbaAnalysis) {
    let Some(procedure) = parse_vba_procedure_name(raw) else {
        return;
    };
    if is_auto_exec_procedure(&procedure) {
        analysis.auto_exec_procedures.push(procedure.clone());
    }
    analysis.procedures.push(procedure);
}

fn suspicious_calls_in_line(raw: &str, line: u32) -> Vec<SuspiciousCall> {
    let lower = raw.to_ascii_lowercase();
    SUSPICIOUS_VBA_CALLS
        .iter()
        .filter(|(pattern, _)| line_contains_pattern(&lower, &pattern.to_ascii_lowercase()))
        .map(|(pattern, category)| SuspiciousCall {
            name: (*pattern).to_string(),
            category: *category,
            line: Some(line),
        })
        .collect()
}

fn line_contains_pattern(line: &str, pattern: &str) -> bool {
    line.match_indices(pattern)
        .any(|(idx, _)| has_identifier_boundary(line, idx, pattern.len()))
}

fn has_identifier_boundary(input: &str, start: usize, len: usize) -> bool {
    is_identifier_boundary(input.as_bytes().get(start.wrapping_sub(1)), start == 0)
        && is_identifier_boundary(
            input.as_bytes().get(start + len),
            start + len >= input.len(),
        )
}

fn is_identifier_boundary(byte: Option<&u8>, at_edge: bool) -> bool {
    at_edge || !byte.is_some_and(|b| b.is_ascii_alphanumeric() || *b == b'_')
}

fn is_comment_line(raw: &str) -> bool {
    raw.starts_with('\'') || raw.starts_with("Rem ") || raw.starts_with("Rem:")
}

fn parse_vba_procedure_name(raw: &str) -> Option<String> {
    let mut tokens = raw.split_whitespace();
    let first = tokens.next()?;
    let keyword = if is_visibility_modifier(first) {
        tokens.next()?
    } else {
        first
    };
    if !(keyword.eq_ignore_ascii_case("sub") || keyword.eq_ignore_ascii_case("function")) {
        return None;
    }

    let name = tokens.next()?.split('(').next().unwrap_or("");
    (!name.is_empty()).then(|| name.to_string())
}

fn is_visibility_modifier(token: &str) -> bool {
    token.eq_ignore_ascii_case("private")
        || token.eq_ignore_ascii_case("public")
        || token.eq_ignore_ascii_case("friend")
        || token.eq_ignore_ascii_case("static")
}

fn formula_contains_xlm_function(formula_upper: &str, function: &str) -> bool {
    let name = function.to_uppercase();
    formula_upper
        .match_indices(&name)
        .any(|(idx, _)| has_formula_boundary(formula_upper, idx, name.len()))
}

fn has_formula_boundary(input: &str, start: usize, len: usize) -> bool {
    formula_boundary_before(input.as_bytes().get(start.wrapping_sub(1)), start == 0)
        && formula_boundary_after(
            input.as_bytes().get(start + len),
            start + len >= input.len(),
        )
}

fn formula_boundary_before(byte: Option<&u8>, at_edge: bool) -> bool {
    at_edge || matches!(byte, Some(b'=' | b'(' | b',' | b' ' | b'\t'))
}

fn formula_boundary_after(byte: Option<&u8>, at_edge: bool) -> bool {
    at_edge || matches!(byte, Some(b'(' | b',' | b')' | b' ' | b'\t' | b'.' | b'!'))
}

#[cfg(test)]
mod tests {
    use super::{
        analyze_vba_source, contains_dangerous_xlm, is_dangerous_xlm_function, AUTO_EXEC_PROCEDURES,
    };
    use docir_core::security::SuspiciousCallCategory;

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

    #[test]
    fn contains_dangerous_xlm_uses_formula_boundaries() {
        assert!(contains_dangerous_xlm("=SUM(A1:A10)").is_empty());
        assert!(!contains_dangerous_xlm("=EXEC(\"cmd /c calc\")").is_empty());
        assert!(!contains_dangerous_xlm("=CALL(\"kernel32\",\"WinExec\")").is_empty());
    }
}
