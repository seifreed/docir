use docir_core::ir::IRNode;
use docir_parser::OoxmlParser;
use std::io::{Cursor, Write};

fn build_docx_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let mut out = Vec::new();
    let cursor = Cursor::new(&mut out);
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::SimpleFileOptions::default();

    // Required OOXML package scaffolding.
    let content_types = br#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/comments.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
          <Override PartName="/docProps/core.xml"
            ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
          <Override PartName="/docProps/app.xml"
            ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
          <Override PartName="/docProps/custom.xml"
            ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
          <Override PartName="/customXml/item1.bin"
            ContentType="application/octet-stream"/>
        </Types>
    "#;
    let package_rels = br#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
          <Relationship Id="rIdCore"
            Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
            Target="docProps/core.xml"/>
          <Relationship Id="rIdApp"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
            Target="docProps/app.xml"/>
          <Relationship Id="rIdCustom"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
            Target="docProps/custom.xml"/>
        </Relationships>
    "#;
    let document_xml = br#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
          </w:body>
        </w:document>
    "#;

    zip.start_file("[Content_Types].xml", options)
        .expect("content types");
    zip.write_all(content_types).expect("content types bytes");
    zip.add_directory("_rels/", options).expect("_rels dir");
    zip.start_file("_rels/.rels", options)
        .expect("package rels");
    zip.write_all(package_rels).expect("package rels bytes");
    zip.add_directory("word/", options).expect("word dir");
    zip.start_file("word/document.xml", options)
        .expect("document part");
    zip.write_all(document_xml).expect("document bytes");

    for (path, bytes) in entries {
        if path.ends_with('/') {
            zip.add_directory(*path, options)
                .expect("extra directory entry");
        } else {
            zip.start_file(path, options).expect("extra part");
            zip.write_all(bytes).expect("extra bytes");
        }
    }

    zip.finish().expect("zip finish");
    out
}

#[test]
fn post_process_only_adds_true_unparsed_parts_and_emits_summary() {
    let core = br#"
        <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                           xmlns:dc="http://purl.org/dc/elements/1.1/">
          <dc:title>Doc</dc:title>
        </cp:coreProperties>
    "#;
    let app = br#"
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
          <Application>docir</Application>
        </Properties>
    "#;
    let custom = br#"
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
                    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>
    "#;
    let comments = br#"
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
    "#;
    let bytes = build_docx_zip(&[
        ("docProps/core.xml", core),
        ("docProps/app.xml", app),
        ("docProps/custom.xml", custom),
        ("word/comments.xml", comments),
        ("customXml/item1.bin", b"BIN"),
        ("[trash]/", b""),
        ("[trash]/skip.bin", b"trash"),
    ]);

    let parser = OoxmlParser::new();
    let parsed = parser
        .parse_reader(Cursor::new(bytes))
        .expect("parse synthetic docx");
    let doc = parsed.document().expect("document");

    let extension_parts: Vec<_> = doc
        .shared_parts
        .iter()
        .filter_map(|id| match parsed.store.get(*id) {
            Some(IRNode::ExtensionPart(part)) => Some(part.path.clone()),
            _ => None,
        })
        .collect();

    assert!(extension_parts.iter().any(|p| p == "customXml/item1.bin"));
    assert!(!extension_parts.iter().any(|p| p == "docProps/core.xml"));
    assert!(!extension_parts.iter().any(|p| p == "docProps/app.xml"));
    assert!(!extension_parts.iter().any(|p| p == "docProps/custom.xml"));
    assert!(!extension_parts.iter().any(|p| p == "word/comments.xml"));
    assert!(!extension_parts.iter().any(|p| p.starts_with("[trash]/")));

    let mut found_unparsed_summary = false;
    let mut found_unparsed_part = false;
    for diag_id in &doc.diagnostics {
        if let Some(IRNode::Diagnostics(diag)) = parsed.store.get(*diag_id) {
            for entry in &diag.entries {
                if entry.code == "UNPARSED_SUMMARY" && entry.message.contains("unparsed parts: 1") {
                    found_unparsed_summary = true;
                }
                if entry.code == "UNPARSED_PART" && entry.message.contains("customXml/item1.bin") {
                    found_unparsed_part = true;
                }
            }
        }
    }
    assert!(found_unparsed_summary);
    assert!(found_unparsed_part);
}
