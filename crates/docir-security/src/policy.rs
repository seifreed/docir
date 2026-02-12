//! Security policy constants and heuristics.

use docir_core::security::SuspiciousCallCategory;

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

/// Known auto-execute macro procedure names.
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
