//! Shared diagnostic helpers.

use docir_core::ir::{DiagnosticEntry, DiagnosticSeverity, Diagnostics};

pub(crate) fn push_entry(
    entries: &mut Vec<DiagnosticEntry>,
    severity: DiagnosticSeverity,
    code: &str,
    message: String,
    path: Option<&str>,
) {
    entries.push(DiagnosticEntry {
        severity,
        code: code.to_string(),
        message,
        path: path.map(|p| p.to_string()),
    });
}

pub(crate) fn push_info(
    diagnostics: &mut Diagnostics,
    code: &str,
    message: String,
    path: Option<&str>,
) {
    push_entry(
        &mut diagnostics.entries,
        DiagnosticSeverity::Info,
        code,
        message,
        path,
    );
}

pub(crate) fn push_warning(
    diagnostics: &mut Diagnostics,
    code: &str,
    message: String,
    path: Option<&str>,
) {
    push_entry(
        &mut diagnostics.entries,
        DiagnosticSeverity::Warning,
        code,
        message,
        path,
    );
}

pub(crate) fn push_error(
    diagnostics: &mut Diagnostics,
    code: &str,
    message: String,
    path: Option<&str>,
) {
    push_entry(
        &mut diagnostics.entries,
        DiagnosticSeverity::Error,
        code,
        message,
        path,
    );
}
