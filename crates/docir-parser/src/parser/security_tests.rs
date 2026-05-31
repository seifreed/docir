use super::*;
use crate::zip_handler::PackageReader;
use docir_core::security::ExternalRefType;
use docir_core::visitor::IrStore;
use std::collections::HashMap;

struct TestPackageReader {
    files: HashMap<String, Vec<u8>>,
}

impl TestPackageReader {
    fn new(entries: &[(&str, &[u8])]) -> Self {
        let files = entries
            .iter()
            .map(|(path, bytes)| ((*path).to_string(), bytes.to_vec()))
            .collect();
        Self { files }
    }
}

impl PackageReader for TestPackageReader {
    fn contains(&self, name: &str) -> bool {
        self.files.contains_key(name)
    }

    fn read_file(&mut self, name: &str) -> Result<Vec<u8>, ParseError> {
        self.files
            .get(name)
            .cloned()
            .ok_or_else(|| ParseError::MissingPart(name.to_string()))
    }

    fn file_size(&mut self, name: &str) -> Result<u64, ParseError> {
        self.files
            .get(name)
            .map(|v| v.len() as u64)
            .ok_or_else(|| ParseError::MissingPart(name.to_string()))
    }

    fn file_names(&self) -> Vec<String> {
        self.files.keys().cloned().collect()
    }

    fn list_prefix(&self, prefix: &str) -> Vec<String> {
        self.files
            .keys()
            .filter(|name| name.starts_with(prefix))
            .cloned()
            .collect()
    }

    fn list_suffix(&self, suffix: &str) -> Vec<String> {
        self.files
            .keys()
            .filter(|name| name.ends_with(suffix))
            .cloned()
            .collect()
    }
}

fn minimal_valid_cfb() -> Vec<u8> {
    const FREE: u32 = 0xFFFF_FFFF;
    const END: u32 = 0xFFFF_FFFE;
    const FAT: u32 = 0xFFFF_FFFD;

    let mut data = vec![0u8; 512 * 3];

    // Header signature.
    data[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    // Sector shift (512-byte sectors), mini sector shift (64-byte sectors).
    data[0x1E..0x20].copy_from_slice(&(9u16).to_le_bytes());
    data[0x20..0x22].copy_from_slice(&(6u16).to_le_bytes());
    // FAT and directory pointers.
    data[0x2C..0x30].copy_from_slice(&(1u32).to_le_bytes()); // num FAT sectors
    data[0x30..0x34].copy_from_slice(&(1u32).to_le_bytes()); // first dir sector
    data[0x38..0x3C].copy_from_slice(&(4096u32).to_le_bytes()); // mini cutoff
    data[0x3C..0x40].copy_from_slice(&END.to_le_bytes()); // first mini FAT
    data[0x40..0x44].copy_from_slice(&(0u32).to_le_bytes()); // num mini FAT sectors
    data[0x44..0x48].copy_from_slice(&END.to_le_bytes()); // first DIFAT sector
    data[0x48..0x4C].copy_from_slice(&(0u32).to_le_bytes()); // num DIFAT sectors
    // DIFAT entries in header: first FAT sector is sector 0.
    data[0x4C..0x50].copy_from_slice(&(0u32).to_le_bytes());
    for idx in 1..109 {
        let off = 0x4C + idx * 4;
        data[off..off + 4].copy_from_slice(&FREE.to_le_bytes());
    }

    // FAT sector (sector 0): [FAT, DIR, FREE...]
    let fat_start = 512;
    data[fat_start..fat_start + 4].copy_from_slice(&FAT.to_le_bytes());
    data[fat_start + 4..fat_start + 8].copy_from_slice(&END.to_le_bytes());
    for idx in 2..128 {
        let off = fat_start + idx * 4;
        data[off..off + 4].copy_from_slice(&FREE.to_le_bytes());
    }

    // Directory sector (sector 1) with only root entry.
    let dir_start = 1024;
    let name_utf16: Vec<u8> = "Root Entry"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .chain([0, 0])
        .collect();
    data[dir_start..dir_start + name_utf16.len()].copy_from_slice(&name_utf16);
    data[dir_start + 64..dir_start + 66].copy_from_slice(&(name_utf16.len() as u16).to_le_bytes());
    data[dir_start + 66] = 5; // Root storage object.
    data[dir_start + 68..dir_start + 72].copy_from_slice(&FREE.to_le_bytes()); // left sibling
    data[dir_start + 72..dir_start + 76].copy_from_slice(&FREE.to_le_bytes()); // right sibling
    data[dir_start + 76..dir_start + 80].copy_from_slice(&FREE.to_le_bytes()); // child
    data[dir_start + 116..dir_start + 120].copy_from_slice(&END.to_le_bytes()); // root stream start
    data[dir_start + 120..dir_start + 128].copy_from_slice(&(0u64).to_le_bytes()); // size

    data
}

#[test]
fn scan_zip_detects_macro_project_without_project_stream() {
    let cfb = minimal_valid_cfb();
    let mut zip = TestPackageReader::new(&[("word/vbaProject.bin", &cfb)]);
    let mut store = IrStore::new();
    let config = ParserConfig::default();
    let scanner = SecurityScanner::new(&config);
    scanner
        .scan_zip(&mut zip, &mut store)
        .expect("scan succeeds");

    let projects: Vec<_> = store
        .values()
        .filter_map(|node| match node {
            IRNode::MacroProject(project) => Some(project),
            _ => None,
        })
        .collect();
    assert_eq!(projects.len(), 1);
    assert_eq!(
        projects[0].span.as_ref().map(|s| s.file_path.as_str()),
        Some("word/vbaProject.bin")
    );
    assert!(projects[0].modules.is_empty());
    assert!(!projects[0].has_auto_exec);
}

#[test]
fn scan_zip_maps_external_relationship_types_and_locations() {
    let rels = br#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rHyper" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test" TargetMode="External"/>
          <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="http://cdn.example.test/a.png" TargetMode="External"/>
          <Relationship Id="rOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="https://example.test/ole" TargetMode="External"/>
          <Relationship Id="rTpl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="https://example.test/template.dotm" TargetMode="External"/>
          <Relationship Id="rOther" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="https://example.test/footer" TargetMode="External"/>
        </Relationships>
    "#;
    let mut zip = TestPackageReader::new(&[("word/_rels/document.xml.rels", rels)]);
    let mut store = IrStore::new();
    let config = ParserConfig::default();
    let scanner = SecurityScanner::new(&config);
    scanner
        .scan_zip(&mut zip, &mut store)
        .expect("scan succeeds");

    let refs: Vec<_> = store
        .values()
        .filter_map(|node| match node {
            IRNode::ExternalReference(ext) => Some(ext),
            _ => None,
        })
        .collect();
    assert_eq!(refs.len(), 5);
    let by_id: HashMap<_, _> = refs
        .iter()
        .filter_map(|r| r.relationship_id.as_deref().map(|id| (id, r.ref_type)))
        .collect();
    assert_eq!(by_id.get("rHyper"), Some(&ExternalRefType::Hyperlink));
    assert_eq!(by_id.get("rImage"), Some(&ExternalRefType::Image));
    assert_eq!(by_id.get("rOle"), Some(&ExternalRefType::OleLink));
    assert_eq!(by_id.get("rTpl"), Some(&ExternalRefType::AttachedTemplate));
    assert_eq!(by_id.get("rOther"), Some(&ExternalRefType::Other));
    assert!(refs.iter().all(|r| r.span.is_some()));
    assert!(
        refs.iter()
            .all(|r| r.relationship_type.as_deref().is_some())
    );
}

#[test]
fn scan_zip_deduplicates_activex_binary_and_scans_other_ole() {
    let activex_xml = br#"<ocx name="Button1" clsid="{ABC}"/>"#;
    let activex_rels = br#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rBin" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>
        </Relationships>
    "#;
    let doc_rels = br#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rLocal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="local.docx" TargetMode="Internal"/>
        </Relationships>
    "#;
    let mut zip = TestPackageReader::new(&[
        ("word/embeddings/object1.bin", b"OLE-1"),
        ("word/activeX/activeX1.xml", activex_xml),
        ("word/activeX/activeX2.xml", activex_xml),
        ("word/activeX/_rels/activeX1.xml.rels", activex_rels),
        ("word/activeX/_rels/activeX2.xml.rels", activex_rels),
        ("word/activeX/activeX1.bin", b"BIN-DATA"),
        ("word/_rels/document.xml.rels", doc_rels),
    ]);
    let mut store = IrStore::new();
    let config = ParserConfig::default();
    let scanner = SecurityScanner::new(&config);
    scanner
        .scan_zip(&mut zip, &mut store)
        .expect("scan succeeds");

    let ole_count = store
        .values()
        .filter(|node| matches!(node, IRNode::OleObject(_)))
        .count();
    let activex_count = store
        .values()
        .filter(|node| matches!(node, IRNode::ActiveXControl(_)))
        .count();
    let external_ref_count = store
        .values()
        .filter(|node| matches!(node, IRNode::ExternalReference(_)))
        .count();

    // One embedding + one deduplicated ActiveX binary.
    assert_eq!(ole_count, 2);
    // Two XML controls still become two control nodes.
    assert_eq!(activex_count, 2);
    // Internal relationship should not create an external reference.
    assert_eq!(external_ref_count, 0);
}

#[test]
fn scan_zip_reports_malformed_activex_xml() {
    let mut zip = TestPackageReader::new(&[(
        "word/activeX/activeX1.xml",
        br#"<ocx name="Button1" clsid="{ABC}"><prop>"#,
    )]);
    let mut store = IrStore::new();
    let config = ParserConfig::default();
    let scanner = SecurityScanner::new(&config);

    let err = scanner
        .scan_zip(&mut zip, &mut store)
        .expect_err("malformed ActiveX XML must fail security scanning");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "word/activeX/activeX1.xml"),
        other => panic!("expected XML error, got {other}"),
    }
}

#[test]
fn scan_zip_sets_ole_hash_when_compute_hashes_enabled() {
    let mut zip = TestPackageReader::new(&[("word/embeddings/object1.bin", b"OLE-HASH")]);
    let mut store = IrStore::new();
    let config = ParserConfig {
        compute_hashes: true,
        ..ParserConfig::default()
    };
    let scanner = SecurityScanner::new(&config);
    scanner
        .scan_zip(&mut zip, &mut store)
        .expect("scan succeeds");

    let ole_nodes: Vec<_> = store
        .values()
        .filter_map(|node| match node {
            IRNode::OleObject(ole) => Some(ole),
            _ => None,
        })
        .collect();
    assert_eq!(ole_nodes.len(), 1);
    assert_eq!(ole_nodes[0].size_bytes, 8);
    assert!(ole_nodes[0].data_hash.is_some());
}

#[test]
fn scan_zip_activex_binary_rel_type_without_bin_extension_is_still_scanned() {
    let activex_xml = br#"<ocx name="ButtonX" clsid="{ABC}"/>"#;
    let activex_rels = br#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rBin" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeXBinary.dat"/>
        </Relationships>
    "#;
    let mut zip = TestPackageReader::new(&[
        ("word/activeX/activeX1.xml", activex_xml),
        ("word/activeX/_rels/activeX1.xml.rels", activex_rels),
        ("word/activeX/activeXBinary.dat", b"BINARY"),
        ("word/_rels/document.xml.rels", b"<Relationships/>"),
    ]);
    let mut store = IrStore::new();
    let config = ParserConfig::default();
    let scanner = SecurityScanner::new(&config);
    scanner
        .scan_zip(&mut zip, &mut store)
        .expect("scan succeeds");

    let ole_count = store
        .values()
        .filter(|node| matches!(node, IRNode::OleObject(_)))
        .count();
    let activex_count = store
        .values()
        .filter(|node| matches!(node, IRNode::ActiveXControl(_)))
        .count();
    assert_eq!(activex_count, 1);
    assert_eq!(ole_count, 1);
}

#[test]
fn scan_zip_collects_external_relationships_from_all_word_rels_files() {
    let rels_main = br#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rMain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/main" TargetMode="External"/>
        </Relationships>
    "#;
    let rels_footnotes = br#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rFoot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/img.png" TargetMode="External"/>
        </Relationships>
    "#;
    let rels_non_word = br#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rSkip" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/skip" TargetMode="External"/>
        </Relationships>
    "#;
    let mut zip = TestPackageReader::new(&[
        ("word/_rels/document.xml.rels", rels_main),
        ("word/_rels/footnotes.xml.rels", rels_footnotes),
        ("xl/_rels/workbook.xml.rels", rels_non_word),
    ]);
    let mut store = IrStore::new();
    let config = ParserConfig::default();
    let scanner = SecurityScanner::new(&config);
    scanner
        .scan_zip(&mut zip, &mut store)
        .expect("scan succeeds");

    let refs: Vec<_> = store
        .values()
        .filter_map(|node| match node {
            IRNode::ExternalReference(ext) => Some(ext),
            _ => None,
        })
        .collect();
    assert_eq!(refs.len(), 2);
    assert!(refs.iter().any(|ext| {
        ext.span
            .as_ref()
            .is_some_and(|s| s.file_path == "word/_rels/document.xml.rels")
    }));
    assert!(refs.iter().any(|ext| {
        ext.span
            .as_ref()
            .is_some_and(|s| s.file_path == "word/_rels/footnotes.xml.rels")
    }));
}
