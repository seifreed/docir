//! # docir-diff
//!
//! Structural and semantic diffing for docir IR.

use docir_core::types::{NodeId, NodeType};
use docir_core::visitor::IrStore;
use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
mod index;
mod summary;

use index::build_index;

/// Diff result between two IR trees.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiffResult {
    /// Nodes only present in the right-hand document.
    pub added: Vec<NodeChange>,
    /// Nodes only present in the left-hand document.
    pub removed: Vec<NodeChange>,
    /// Nodes present in both but with differing summaries.
    pub modified: Vec<NodeModification>,
}

impl DiffResult {
    /// Returns true if there are no differences.
    pub fn is_empty(&self) -> bool {
        self.added.is_empty() && self.removed.is_empty() && self.modified.is_empty()
    }
}

/// A node that was added or removed.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NodeChange {
    pub key: String,
    pub node_type: NodeType,
    pub summary: String,
}

/// A node that was modified.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NodeModification {
    pub key: String,
    pub node_type: NodeType,
    pub before: String,
    pub after: String,
    pub change_kind: ChangeKind,
}

/// Kind of change detected between two matched nodes.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum ChangeKind {
    Content,
    Style,
    Both,
    Metadata,
}

/// Diff engine for docir IR.
pub struct DiffEngine;

impl DiffEngine {
    /// Computes a diff between two IR trees.
    pub fn diff(
        left: &IrStore,
        left_root: NodeId,
        right: &IrStore,
        right_root: NodeId,
    ) -> DiffResult {
        let left_index = build_index(left, left_root);
        let right_index = build_index(right, right_root);

        let mut added = Vec::new();
        let mut removed = Vec::new();
        let mut modified = Vec::new();

        let mut keys: BTreeMap<&String, ()> = BTreeMap::new();
        for key in left_index.keys() {
            keys.insert(key, ());
        }
        for key in right_index.keys() {
            keys.insert(key, ());
        }

        for key in keys.keys() {
            match (left_index.get(*key), right_index.get(*key)) {
                (Some(left_snap), Some(right_snap)) => {
                    if left_snap.summary != right_snap.summary {
                        let content_same = left_snap.content_sig == right_snap.content_sig;
                        let style_same = left_snap.style_sig == right_snap.style_sig;
                        let change_kind = if content_same && !style_same {
                            ChangeKind::Style
                        } else if !content_same && style_same {
                            ChangeKind::Content
                        } else if !content_same && !style_same {
                            ChangeKind::Both
                        } else {
                            ChangeKind::Metadata
                        };
                        modified.push(NodeModification {
                            key: (*key).clone(),
                            node_type: left_snap.node_type,
                            before: left_snap.summary.clone(),
                            after: right_snap.summary.clone(),
                            change_kind,
                        });
                    }
                }
                (Some(left_snap), None) => {
                    removed.push(NodeChange {
                        key: (*key).clone(),
                        node_type: left_snap.node_type,
                        summary: left_snap.summary.clone(),
                    });
                }
                (None, Some(right_snap)) => {
                    added.push(NodeChange {
                        key: (*key).clone(),
                        node_type: right_snap.node_type,
                        summary: right_snap.summary.clone(),
                    });
                }
                (None, None) => {}
            }
        }

        DiffResult {
            added,
            removed,
            modified,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_parser::DocumentParser;
    use std::io::{Cursor, Write};
    use zip::write::FileOptions;

    fn build_odf_zip(content_xml: &str) -> Vec<u8> {
        let mut buffer = Vec::new();
        let cursor = Cursor::new(&mut buffer);
        let mut zip = zip::ZipWriter::new(cursor);
        let stored =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
        zip.start_file("mimetype", stored).unwrap();
        zip.write_all(b"application/vnd.oasis.opendocument.text")
            .unwrap();

        zip.start_file("META-INF/manifest.xml", FileOptions::<()>::default())
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#,
        )
        .unwrap();

        zip.start_file("content.xml", FileOptions::<()>::default())
            .unwrap();
        zip.write_all(content_xml.as_bytes()).unwrap();

        zip.start_file("meta.xml", FileOptions::<()>::default())
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
  <office:meta>
    <dc:title>Diff</dc:title>
  </office:meta>
</office:document-meta>
"#,
        )
        .unwrap();

        zip.finish().unwrap();
        buffer
    }

    #[test]
    fn test_diff_odt_modified_text() {
        let content_left = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:p>Hello</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
        let content_right = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:p>Hello world</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
        let left_zip = build_odf_zip(content_left);
        let right_zip = build_odf_zip(content_right);

        let parser = DocumentParser::new();
        let left_doc = parser.parse_reader(Cursor::new(left_zip)).unwrap();
        let right_doc = parser.parse_reader(Cursor::new(right_zip)).unwrap();

        let diff = DiffEngine::diff(
            &left_doc.store,
            left_doc.root_id,
            &right_doc.store,
            right_doc.root_id,
        );

        assert!(diff
            .modified
            .iter()
            .any(|m| m.change_kind == ChangeKind::Content));
    }
}
