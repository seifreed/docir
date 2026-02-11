//! Shared diagnostic helpers.

use docir_core::ir::{DiagnosticEntry, DiagnosticSeverity, Diagnostics};

pub(crate) fn push_warning(
    diagnostics: &mut Diagnostics,
    code: &str,
    message: String,
    path: Option<&str>,
) {
    diagnostics.entries.push(DiagnosticEntry {
        severity: DiagnosticSeverity::Warning,
        code: code.to_string(),
        message,
        path: path.map(|p| p.to_string()),
    });
}
