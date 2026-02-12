use crate::ir::{IRNode, IrNode};
use crate::types::NodeId;
use std::collections::HashMap;

/// Storage for IR nodes indexed by NodeId.
#[derive(Debug, Clone, Default)]
pub struct IrStore {
    nodes: HashMap<NodeId, IRNode>,
}

impl IrStore {
    /// Creates a new empty IrStore.
    pub fn new() -> Self {
        Self {
            nodes: HashMap::new(),
        }
    }

    /// Consumes the store and returns all nodes.
    pub fn into_nodes(self) -> Vec<IRNode> {
        self.nodes.into_values().collect()
    }

    /// Inserts a node into the store.
    pub fn insert(&mut self, node: IRNode) {
        let id = node.node_id();
        self.nodes.insert(id, node);
    }

    /// Gets a node by ID.
    pub fn get(&self, id: NodeId) -> Option<&IRNode> {
        self.nodes.get(&id)
    }

    /// Gets a mutable reference to a node by ID.
    pub fn get_mut(&mut self, id: NodeId) -> Option<&mut IRNode> {
        self.nodes.get_mut(&id)
    }

    /// Returns the number of nodes.
    pub fn len(&self) -> usize {
        self.nodes.len()
    }

    /// Returns node IDs by type.
    pub fn iter_ids_by_type(
        &self,
        node_type: crate::types::NodeType,
    ) -> impl Iterator<Item = NodeId> + '_ {
        self.nodes.iter().filter_map(move |(id, node)| {
            if node.node_type() == node_type {
                Some(*id)
            } else {
                None
            }
        })
    }

    /// Returns an iterator over all nodes.
    pub fn values(&self) -> impl Iterator<Item = &IRNode> + '_ {
        self.nodes.values()
    }

    /// Returns true if the store is empty.
    pub fn is_empty(&self) -> bool {
        self.nodes.is_empty()
    }

    /// Extends the store with nodes from another store.
    pub fn extend(&mut self, other: IrStore) {
        for (id, node) in other.nodes {
            self.nodes.insert(id, node);
        }
    }

    /// Iterates over all nodes.
    pub fn iter(&self) -> impl Iterator<Item = (&NodeId, &IRNode)> {
        self.nodes.iter()
    }
}
