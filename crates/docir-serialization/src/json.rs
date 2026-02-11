//! JSON serialization for IR.

use crate::{IrSerializer, SerializationError};
use docir_core::ir::{IRNode, IrNode};
use docir_core::visitor::IrStore;
use docir_core::NodeId;
use serde::Serialize;
use serde_json::Value;
use std::collections::BTreeMap;

/// JSON serializer configuration.
#[derive(Debug, Clone)]
pub struct JsonSerializer {
    /// Pretty-print output.
    pub pretty: bool,
    /// Include source span information.
    pub include_spans: bool,
    /// Sort object keys for deterministic output.
    pub sort_keys: bool,
}

impl Default for JsonSerializer {
    fn default() -> Self {
        Self {
            pretty: false,
            include_spans: true,
            sort_keys: true,
        }
    }
}

impl JsonSerializer {
    /// Creates a new JSON serializer.
    pub fn new() -> Self {
        Self::default()
    }

    /// Creates a pretty-printing JSON serializer.
    pub fn pretty() -> Self {
        Self {
            pretty: true,
            ..Default::default()
        }
    }

    /// Builds a complete IR tree as a serializable structure.
    fn build_tree(&self, store: &IrStore, root_id: NodeId) -> Result<TreeNode, SerializationError> {
        let node = store
            .get(root_id)
            .ok_or_else(|| SerializationError::NodeNotFound(format!("{}", root_id)))?;

        let children: Vec<TreeNode> = node
            .children()
            .into_iter()
            .filter_map(|child_id| self.build_tree(store, child_id).ok())
            .collect();

        Ok(TreeNode {
            node: node.clone(),
            children,
        })
    }
}

impl IrSerializer for JsonSerializer {
    fn serialize_node(&self, node: &IRNode) -> Result<Vec<u8>, SerializationError> {
        let mut value = serde_json::to_value(node)?;
        if self.sort_keys {
            sort_json_value(&mut value);
        }
        if self.pretty {
            Ok(serde_json::to_vec_pretty(&value)?)
        } else {
            Ok(serde_json::to_vec(&value)?)
        }
    }

    fn serialize_tree(
        &self,
        store: &IrStore,
        root_id: NodeId,
    ) -> Result<Vec<u8>, SerializationError> {
        let tree = self.build_tree(store, root_id)?;
        let mut value = serde_json::to_value(tree)?;
        if self.sort_keys {
            sort_json_value(&mut value);
        }
        if self.pretty {
            Ok(serde_json::to_vec_pretty(&value)?)
        } else {
            Ok(serde_json::to_vec(&value)?)
        }
    }

    fn serialize_to_string(
        &self,
        store: &IrStore,
        root_id: NodeId,
    ) -> Result<String, SerializationError> {
        let tree = self.build_tree(store, root_id)?;
        let mut value = serde_json::to_value(tree)?;
        if self.sort_keys {
            sort_json_value(&mut value);
        }
        if self.pretty {
            Ok(serde_json::to_string_pretty(&value)?)
        } else {
            Ok(serde_json::to_string(&value)?)
        }
    }
}

fn sort_json_value(value: &mut Value) {
    match value {
        Value::Object(map) => {
            let mut ordered = BTreeMap::new();
            let drained = std::mem::take(map);
            for (k, mut v) in drained {
                sort_json_value(&mut v);
                ordered.insert(k, v);
            }
            for (k, v) in ordered {
                map.insert(k, v);
            }
        }
        Value::Array(items) => {
            for item in items {
                sort_json_value(item);
            }
        }
        _ => {}
    }
}

/// A tree node for serialization (node + children).
#[derive(Debug, Clone, Serialize)]
struct TreeNode {
    #[serde(flatten)]
    node: IRNode,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    children: Vec<TreeNode>,
}

/// Serializes a parsed document to JSON.
pub fn to_json(
    store: &IrStore,
    root_id: NodeId,
    pretty: bool,
) -> Result<String, SerializationError> {
    let serializer = if pretty {
        JsonSerializer::pretty()
    } else {
        JsonSerializer::new()
    };

    serializer.serialize_to_string(store, root_id)
}

/// Serializes just the flat node store to JSON.
pub fn store_to_json(store: &IrStore, pretty: bool) -> Result<String, SerializationError> {
    // Collect nodes into a sorted map for deterministic output
    let nodes: BTreeMap<String, &IRNode> = store
        .iter()
        .map(|(id, node)| (format!("{}", id), node))
        .collect();

    if pretty {
        Ok(serde_json::to_string_pretty(&nodes)?)
    } else {
        Ok(serde_json::to_string(&nodes)?)
    }
}
