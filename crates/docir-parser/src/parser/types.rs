use super::*;

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
