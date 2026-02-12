//! Security policy constants and heuristics.

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
    // PowerPoint
    "Auto_Open",
    "Auto_Close",
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

pub fn is_dangerous_xlm_function(name: &str) -> bool {
    let upper = name.to_ascii_uppercase();
    DANGEROUS_XLM_FUNCTIONS.iter().any(|&item| item == upper)
}

pub fn analyze_vba_source(source: &str) -> VbaAnalysis {
    let mut analysis = VbaAnalysis::default();

    for (idx, line) in source.lines().enumerate() {
        let raw = line.trim();
        if raw.is_empty() {
            continue;
        }

        let lower = raw.to_ascii_lowercase();
        let mut tokens: Vec<&str> = raw.split_whitespace().collect();
        if tokens.len() >= 2 {
            if tokens[0].eq_ignore_ascii_case("private")
                || tokens[0].eq_ignore_ascii_case("public")
                || tokens[0].eq_ignore_ascii_case("friend")
                || tokens[0].eq_ignore_ascii_case("static")
            {
                tokens.remove(0);
            }
        }
        if tokens.len() >= 2 {
            let keyword = tokens[0].to_ascii_lowercase();
            if keyword == "sub" || keyword == "function" {
                let name = tokens[1].split('(').next().unwrap_or("").to_string();
                if !name.is_empty() {
                    analysis.procedures.push(name);
                }
            }
        }

        for (pattern, category) in SUSPICIOUS_VBA_CALLS {
            if lower.contains(&pattern.to_ascii_lowercase()) {
                analysis.suspicious_calls.push(SuspiciousCall {
                    name: (*pattern).to_string(),
                    category: *category,
                    line: Some(idx as u32 + 1),
                });
            }
        }

        for proc in AUTO_EXEC_PROCEDURES {
            if lower.starts_with("sub ") && lower.contains(&proc.to_ascii_lowercase()) {
                analysis.auto_exec_procedures.push(proc.to_string());
            }
        }
    }

    analysis
}
