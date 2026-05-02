//! # docir-diff
//!
//! Structural and semantic diffing for docir IR.

use docir_core::types::{NodeId, NodeType};
use docir_core::visitor::IrStore;
use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
use thiserror::Error;

mod index;
mod summary;

use index::build_index;

/// Result of a differential comparison between two IR trees.
///
/// The payload is grouped by change shape:
/// - `added`: nodes only in the right-hand document.
/// - `removed`: nodes only in the left-hand document.
/// - `modified`: nodes present in both documents with changed summaries.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiffResult {
    /// Nodes only present in the right-hand document.
    pub added: Vec<NodeChange>,
    /// Nodes only present in the left-hand document.
    pub removed: Vec<NodeChange>,
    /// Nodes present in both documents but with differing summaries.
    pub modified: Vec<NodeModification>,
}

impl DiffResult {
    /// Returns true when no diff has been detected.
    pub fn is_empty(&self) -> bool {
        self.added.is_empty() && self.removed.is_empty() && self.modified.is_empty()
    }
}

/// Diff-specific error type.
#[derive(Debug, Error)]
pub enum DiffError {
    /// Could not build an index for one of the input stores.
    #[error("diff node indexing failed: {0}")]
    NodeIndexing(String),
    /// Could not build a test ZIP for parser-driven diff fixtures.
    #[error("diff test zip build failed: {0}")]
    TestZipBuild(String),
}

/// Node that was added or removed.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NodeChange {
    /// Stable key used by the IR index and diff algorithm.
    pub key: String,
    /// Node type for the changed element.
    pub node_type: NodeType,
    /// Human summary shown in report output.
    pub summary: String,
}

/// Node that was modified between two IR snapshots.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NodeModification {
    /// Stable key used by the IR index and diff algorithm.
    pub key: String,
    /// Node type for the changed element.
    pub node_type: NodeType,
    /// Summary text before applying the change.
    pub before: String,
    /// Summary text after applying the change.
    pub after: String,
    /// Change classification inferred by signature comparison.
    pub change_kind: ChangeKind,
}

/// Change classification for a pair of compared nodes.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum ChangeKind {
    /// Only content changed.
    Content,
    /// Only style metadata changed.
    Style,
    /// Both content and style changed.
    Both,
    /// Change is limited to metadata.
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
    ) -> Result<DiffResult, DiffError> {
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
                (None, None) => unreachable!("key must exist in at least one index"),
            }
        }

        Ok(DiffResult {
            added,
            removed,
            modified,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_parser::DocumentParser;
    use std::io::{Cursor, Write};
    use zip::write::FileOptions;

    fn build_odf_zip(content_xml: &str) -> Result<Vec<u8>, DiffError> {
        let mut buffer = Vec::new();
        let cursor = Cursor::new(&mut buffer);
        let mut zip = zip::ZipWriter::new(cursor);
        let stored =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);

        zip.start_file("mimetype", stored)
            .map_err(|err| map_zip_error("start mimetype file", err))?;
        write_exact(
            &mut zip,
            b"application/vnd.oasis.opendocument.text",
            "write mimetype payload",
        )?;
        zip.start_file("META-INF/manifest.xml", FileOptions::<()>::default())
            .map_err(|err| map_zip_error("start manifest file", err))?;
        write_exact(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#,
            "write manifest xml",
        )?;

        zip.start_file("content.xml", FileOptions::<()>::default())
            .map_err(|err| map_zip_error("start content file", err))?;
        write_exact(&mut zip, content_xml.as_bytes(), "write content xml")?;

        zip.start_file("meta.xml", FileOptions::<()>::default())
            .map_err(|err| map_zip_error("start meta file", err))?;
        write_exact(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
  <office:meta>
    <dc:title>Diff</dc:title>
  </office:meta>
</office:document-meta>
"#,
            "write meta xml",
        )?;

        zip.start_file("styles.xml", FileOptions::<()>::default())
            .map_err(|err| map_zip_error("start styles file", err))?;
        write_exact(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>
"#,
            "write styles xml",
        )?;

        zip.finish()
            .map_err(|err| map_zip_error("finish zip", err))?;
        Ok(buffer)
    }

    fn map_zip_error(operation: &'static str, err: zip::result::ZipError) -> DiffError {
        DiffError::TestZipBuild(format!("{operation}: {err}"))
    }

    fn write_exact<W: Write>(
        target: &mut W,
        data: &[u8],
        operation: &'static str,
    ) -> Result<(), DiffError> {
        target
            .write_all(data)
            .map_err(|err| io_error(operation, err))
    }

    fn io_error(operation: &'static str, err: std::io::Error) -> DiffError {
        DiffError::TestZipBuild(format!("{operation}: {err}"))
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
        let left_zip = build_odf_zip(content_left).expect("zip fixture");
        let right_zip = build_odf_zip(content_right).expect("zip fixture");

        let parser = DocumentParser::new();
        let left_doc = parser.parse_reader(Cursor::new(left_zip)).unwrap();
        let right_doc = parser.parse_reader(Cursor::new(right_zip)).unwrap();

        let diff = DiffEngine::diff(
            &left_doc.store,
            left_doc.root_id,
            &right_doc.store,
            right_doc.root_id,
        )
        .expect("diff should succeed");

        assert!(diff
            .modified
            .iter()
            .any(|m| m.change_kind == ChangeKind::Content));
    }
}
