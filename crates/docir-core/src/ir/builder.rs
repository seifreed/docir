use crate::ir::{IRNode, IrNode as IrNodeTrait};
use crate::types::NodeId;
use crate::visitor::IrStore;

/// Helper for inserting IR nodes with a consistent API.
pub struct IrBuilder<'a> {
    store: &'a mut IrStore,
}

impl<'a> IrBuilder<'a> {
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
