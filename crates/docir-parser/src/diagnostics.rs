//! Shared diagnostic helpers.

use docir_core::ir::{DiagnosticEntry, DiagnosticSeverity, Diagnostics, Document, IRNode};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;

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

pub(crate) fn attach_diagnostics_if_any(
    store: &mut IrStore,
    doc: &mut Document,
    diagnostics: Diagnostics,
) {
    if diagnostics.entries.is_empty() {
        return;
    }

    let diag_id = diagnostics.id;
    store.insert(IRNode::Diagnostics(diagnostics));
    doc.add_diagnostic(diag_id);
}

pub(crate) fn attach_diagnostics_if_any_to_store(
    store: &mut IrStore,
    root_id: NodeId,
    diagnostics: Diagnostics,
) {
    if diagnostics.entries.is_empty() {
        return;
    }

    let diag_id = diagnostics.id;
    store.insert(IRNode::Diagnostics(diagnostics));
    if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
        doc.add_diagnostic(diag_id);
    }
}
