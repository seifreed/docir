use super::ParseMetrics;
use docir_core::ir::{Document, IRNode};
use docir_core::security::SecurityInfo;
use docir_core::types::{DocumentFormat, NodeId};
use docir_core::visitor::IrStore;

/// Parsed document result.
#[derive(Debug)]
pub struct ParsedDocument {
    /// Root document node ID.
    pub root_id: NodeId,
    /// Document format.
    pub format: DocumentFormat,
    /// IR node store.
    pub store: IrStore,
    /// Optional parse metrics.
    pub metrics: Option<ParseMetrics>,
}

impl ParsedDocument {
    /// Gets the root document node.
    pub fn document(&self) -> Option<&Document> {
        if let Some(IRNode::Document(doc)) = self.store.get(self.root_id) {
            Some(doc)
        } else {
            None
        }
    }

    /// Gets the security info from the document.
    pub fn security_info(&self) -> Option<&SecurityInfo> {
        self.document().map(|d| &d.security)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::Document;
    use docir_core::visitor::IrStore;

    #[test]
    fn document_and_security_info_follow_root_document_node() {
        let mut store = IrStore::new();
        let doc = Document::new(DocumentFormat::WordProcessing);
        let root = doc.id;
        store.insert(IRNode::Document(doc));

        let parsed = ParsedDocument {
            root_id: root,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        };

        assert!(parsed.document().is_some());
        assert!(parsed.security_info().is_some());
    }

    #[test]
    fn document_returns_none_for_non_document_root() {
        let mut store = IrStore::new();
        let para = docir_core::ir::Paragraph::new();
        let root = para.id;
        store.insert(IRNode::Paragraph(para));

        let parsed = ParsedDocument {
            root_id: root,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        };

        assert!(parsed.document().is_none());
        assert!(parsed.security_info().is_none());
    }
}
