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
    fn build_tree(store: &IrStore, root_id: NodeId) -> Result<TreeNode, SerializationError> {
        let node = store
            .get(root_id)
            .ok_or_else(|| SerializationError::NodeNotFound(format!("{}", root_id)))?;

        let children: Vec<TreeNode> = node
            .children()
            .into_iter()
            .filter_map(|child_id| Self::build_tree(store, child_id).ok())
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
        if !self.include_spans {
            strip_spans(&mut value);
        }
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
        let tree = Self::build_tree(store, root_id)?;
        let mut value = serde_json::to_value(tree)?;
        if !self.include_spans {
            strip_spans(&mut value);
        }
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
        let tree = Self::build_tree(store, root_id)?;
        let mut value = serde_json::to_value(tree)?;
        if !self.include_spans {
            strip_spans(&mut value);
        }
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

fn strip_spans(value: &mut Value) {
    match value {
        Value::Object(map) => {
            map.remove("span");
            for nested in map.values_mut() {
                strip_spans(nested);
            }
        }
        Value::Array(items) => {
            for item in items {
                strip_spans(item);
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

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::{Document, IRNode};
    use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
    use docir_core::visitor::IrStore;

    fn sample_store_with_root() -> (IrStore, NodeId) {
        let mut store = IrStore::new();
        let mut doc = Document::new(DocumentFormat::WordProcessing);
        doc.id = NodeId::from_raw(42);
        doc.span = Some(SourceSpan::new("word/document.xml"));
        let root = doc.id;
        store.insert(IRNode::Document(doc));
        (store, root)
    }

    #[test]
    fn to_json_pretty_and_compact_are_semantically_equal() {
        let (store, root) = sample_store_with_root();

        let compact = to_json(&store, root, false).expect("compact JSON");
        let pretty = to_json(&store, root, true).expect("pretty JSON");

        assert!(!compact.contains('\n'));
        assert!(pretty.contains('\n'));
        let compact_val: Value = serde_json::from_str(&compact).expect("valid compact JSON");
        let pretty_val: Value = serde_json::from_str(&pretty).expect("valid pretty JSON");
        assert_eq!(compact_val, pretty_val);
    }

    #[test]
    fn serializer_can_include_or_exclude_spans() {
        let (store, root) = sample_store_with_root();

        let include = JsonSerializer::new()
            .serialize_to_string(&store, root)
            .expect("serialization with spans");
        assert!(include.contains("\"span\""));

        let serializer = JsonSerializer {
            include_spans: false,
            ..JsonSerializer::new()
        };
        let exclude = serializer
            .serialize_to_string(&store, root)
            .expect("serialization without spans");
        assert!(!exclude.contains("\"span\""));
    }

    #[test]
    fn serializer_orders_keys_deterministically() {
        let mut doc = Document::new(DocumentFormat::WordProcessing);
        doc.id = NodeId::from_raw(11);
        let node = IRNode::Document(doc);
        let bytes = JsonSerializer::new()
            .serialize_node(&node)
            .expect("node serialization");
        let json = String::from_utf8(bytes).expect("utf-8 JSON");

        let comments = json.find("\"comments\"").expect("comments key");
        let content = json.find("\"content\"").expect("content key");
        let id = json.find("\"id\"").expect("id key");
        let security = json.find("\"security\"").expect("security key");
        assert!(comments < content);
        assert!(id < security);
    }

    #[test]
    fn to_json_returns_node_not_found_for_unknown_root() {
        let store = IrStore::new();
        let err = to_json(&store, NodeId::from_raw(999_999), false).expect_err("missing root");
        match err {
            SerializationError::NodeNotFound(id) => assert!(id.contains("node_")),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn store_to_json_is_deterministic_and_parses() {
        let (store, _) = sample_store_with_root();
        let compact = store_to_json(&store, false).expect("store compact JSON");
        let pretty = store_to_json(&store, true).expect("store pretty JSON");
        assert!(!compact.is_empty());
        assert!(pretty.contains('\n'));

        let compact_val: Value = serde_json::from_str(&compact).expect("valid compact JSON");
        let pretty_val: Value = serde_json::from_str(&pretty).expect("valid pretty JSON");
        assert_eq!(compact_val, pretty_val);
    }
}
