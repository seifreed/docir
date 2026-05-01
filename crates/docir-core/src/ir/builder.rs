use crate::ir::{IRNode, IrNode as IrNodeTrait};
use crate::types::NodeId;
use crate::visitor::IrStore;

/// Helper for inserting IR nodes with a consistent API.
pub struct IrBuilder<'a> {
    store: &'a mut IrStore,
}

impl<'a> IrBuilder<'a> {
    /// Public API entrypoint: new.
    pub fn new(store: &'a mut IrStore) -> Self {
        Self { store }
    }

    pub fn insert(&mut self, node: IRNode) -> NodeId {
        let id = node.node_id();
        self.store.insert(node);
        id
    }

    pub fn store(&self) -> &IrStore {
        self.store
    }

    pub fn store_mut(&mut self) -> &mut IrStore {
        self.store
    }
}

#[cfg(test)]
mod tests {
    use super::IrBuilder;
    use crate::ir::{Document, IRNode, IrNode as IrNodeTrait, Paragraph};
    use crate::types::DocumentFormat;
    use crate::visitor::IrStore;

    #[test]
    fn insert_returns_node_id_and_stores_node() {
        let mut store = IrStore::new();
        let mut builder = IrBuilder::new(&mut store);

        let paragraph = Paragraph::new();
        let expected_id = paragraph.id;
        let inserted_id = builder.insert(IRNode::Paragraph(paragraph));

        assert_eq!(inserted_id, expected_id);
        assert!(builder.store().get(inserted_id).is_some());
        assert_eq!(builder.store().len(), 1);
    }

    #[test]
    fn store_mut_allows_direct_mutation() {
        let mut store = IrStore::new();
        let mut builder = IrBuilder::new(&mut store);
        let document = Document::new(DocumentFormat::WordProcessing);
        let document_id = document.id;
        builder.insert(IRNode::Document(document));

        {
            let store_mut = builder.store_mut();
            let node = store_mut
                .get_mut(document_id)
                .expect("document node should exist");
            let IRNode::Document(document) = node else {
                panic!("expected document node");
            };
            document.content.push(document_id);
        }

        let stored = builder
            .store()
            .get(document_id)
            .expect("document node should be retrievable");
        assert_eq!(stored.node_id(), document_id);
        let IRNode::Document(document) = stored else {
            panic!("expected document node");
        };
        assert_eq!(document.content.len(), 1);
    }
}
