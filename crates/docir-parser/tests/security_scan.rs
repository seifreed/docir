use docir_core::ir::IRNode;
use docir_core::security::ExternalRefType;
use docir_core::visitor::IrStore;
use docir_parser::{scan_security_bytes, ParseError, ParserConfig};
use std::io::{Cursor, Write};

fn build_zip(entries: &[(&str, &[u8])]) -> Result<Vec<u8>, ParseError> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::SimpleFileOptions::default();
    for (path, data) in entries {
        zip.start_file(path, options)
            .map_err(|e| ParseError::InvalidZip(format!("zip start_file failed: {e}")))?;
        zip.write_all(data)
            .map_err(|e| ParseError::InvalidZip(format!("zip write_all failed: {e}")))?;
    }
    zip.finish()
        .map_err(|e| ParseError::InvalidZip(format!("zip finish failed: {e}")))
        .map(|c| c.into_inner())
}

#[test]
fn scan_security_bytes_detects_external_refs_ole_and_activex() {
    let activex_xml = br#"<ocx name="Button1" clsid="{ABCDEF-1234}"/>"#;
    let activex_rels = br#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rBin" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>
        </Relationships>
    "#;
    let doc_rels = br#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rHyper" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test" TargetMode="External"/>
          <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://cdn.example.test/image.png" TargetMode="External"/>
          <Relationship Id="rOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="https://example.test/ole" TargetMode="External"/>
          <Relationship Id="rTpl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="https://example.test/template.dotm" TargetMode="External"/>
          <Relationship Id="rOther" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="https://example.test/footer" TargetMode="External"/>
          <Relationship Id="rInternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="local.docx" TargetMode="Internal"/>
        </Relationships>
    "#;

    let zip_bytes = build_zip(&[
        ("word/embeddings/object1.bin", b"OLE-1"),
        ("xl/embeddings/object2.ole", b"OLE-2"),
        ("word/activeX/activeX1.xml", activex_xml),
        ("word/activeX/activeX2.xml", activex_xml),
        ("word/activeX/_rels/activeX1.xml.rels", activex_rels),
        ("word/activeX/_rels/activeX2.xml.rels", activex_rels),
        ("word/activeX/activeX1.bin", b"BIN-DATA"),
        ("word/_rels/document.xml.rels", doc_rels),
    ])
    .expect("zip");

    let mut store = IrStore::new();
    let config = ParserConfig::default();
    scan_security_bytes(&config, &zip_bytes, &mut store).expect("scan security");

    let ole_nodes: Vec<_> = store
        .values()
        .filter_map(|node| match node {
            IRNode::OleObject(ole) => Some(ole),
            _ => None,
        })
        .collect();
    let active_x_nodes = store
        .values()
        .filter(|node| matches!(node, IRNode::ActiveXControl(_)))
        .count();
    let external_refs: Vec<_> = store
        .values()
        .filter_map(|node| match node {
            IRNode::ExternalReference(ext) => Some(ext),
            _ => None,
        })
        .collect();

    assert_eq!(
        ole_nodes.len(),
        3,
        "2 embeddings + 1 deduplicated activeX binary"
    );
    assert_eq!(active_x_nodes, 2);
    assert_eq!(external_refs.len(), 5);
    assert!(ole_nodes.iter().all(|ole| ole.data_hash.is_some()));

    let mut types = external_refs.iter().map(|r| r.ref_type).collect::<Vec<_>>();
    types.sort_by_key(|ty| format!("{ty:?}"));
    assert!(types.contains(&ExternalRefType::Hyperlink));
    assert!(types.contains(&ExternalRefType::Image));
    assert!(types.contains(&ExternalRefType::OleLink));
    assert!(types.contains(&ExternalRefType::AttachedTemplate));
    assert!(types.contains(&ExternalRefType::Other));
}

#[test]
fn scan_security_bytes_respects_compute_hashes_flag_for_ole_objects() {
    let zip_bytes = build_zip(&[("word/embeddings/object1.bin", b"OLE-RAW")]).expect("zip");
    let mut store = IrStore::new();
    let config = ParserConfig {
        compute_hashes: false,
        ..ParserConfig::default()
    };
    scan_security_bytes(&config, &zip_bytes, &mut store).expect("scan security");

    let ole_nodes: Vec<_> = store
        .values()
        .filter_map(|node| match node {
            IRNode::OleObject(ole) => Some(ole),
            _ => None,
        })
        .collect();
    assert_eq!(ole_nodes.len(), 1);
    assert!(ole_nodes[0].data_hash.is_none());
}
