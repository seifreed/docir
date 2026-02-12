use crate::parser::ParsedDocument;
use docir_core::ir::{Diagnostics, Document, IRNode};
use docir_core::normalize::normalize_store;
use docir_core::types::DocumentFormat;
use docir_core::visitor::IrStore;

pub(crate) fn init_document_state(format: DocumentFormat) -> (IrStore, Document, Diagnostics) {
    let store = IrStore::new();
    let doc = Document::new(format);
    let diagnostics = Diagnostics::new();
    (store, doc, diagnostics)
}

pub(crate) fn init_store_and_document(format: DocumentFormat) -> (IrStore, Document) {
    let store = IrStore::new();
    let doc = Document::new(format);
    (store, doc)
}

pub(crate) fn finalize_document(
    format: DocumentFormat,
    mut store: IrStore,
    doc: Document,
) -> ParsedDocument {
    let doc_id = doc.id;
    store.insert(IRNode::Document(doc));
    ParsedDocument {
        root_id: doc_id,
        format,
        store,
        metrics: None,
    }
}

pub(crate) fn finalize_and_normalize(
    format: DocumentFormat,
    store: IrStore,
    doc: Document,
) -> ParsedDocument {
    let mut parsed = finalize_document(format, store, doc);
    normalize_store(&mut parsed.store, parsed.root_id);
    parsed
}
