use std::fs::File;
use std::io::Write;
use std::time::{SystemTime, UNIX_EPOCH};
use zip::write::FileOptions;

pub(super) fn create_minimal_docx(include_document: bool) -> std::path::PathBuf {
    let mut path = std::env::temp_dir();
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    path.push(format!("docir_minimal_{nanos}.docx"));

    let file = File::create(&path).expect("create temp docx");
    let mut zip = zip::ZipWriter::new(file);
    let options = FileOptions::<()>::default();

    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>"#;

    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;

    let document = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
          </w:body>
        </w:document>"#;

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.trim().as_bytes()).unwrap();
    zip.add_directory("_rels/", options).unwrap();
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(rels.trim().as_bytes()).unwrap();
    zip.add_directory("word/", options).unwrap();
    if include_document {
        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(document.trim().as_bytes()).unwrap();
    }
    zip.finish().unwrap();
    path
}

pub(super) fn create_docx_with_body(body_xml: &str) -> std::path::PathBuf {
    let mut path = std::env::temp_dir();
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    path.push(format!("docir_body_{nanos}.docx"));

    let file = File::create(&path).expect("create temp docx");
    let mut zip = zip::ZipWriter::new(file);
    let options = FileOptions::<()>::default();

    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>"#;

    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.trim().as_bytes()).unwrap();
    zip.add_directory("_rels/", options).unwrap();
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(rels.trim().as_bytes()).unwrap();
    zip.add_directory("word/", options).unwrap();
    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(body_xml.trim().as_bytes()).unwrap();
    zip.finish().unwrap();
    path
}

pub(super) fn create_docx_with_relationships(
    body_xml: &str,
    rels_xml: &str,
    content_types: &str,
    extra_files: &[(&str, &str)],
) -> std::path::PathBuf {
    let mut path = std::env::temp_dir();
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    path.push(format!("docir_docx_{nanos}.docx"));

    let file = File::create(&path).expect("create temp docx");
    let mut zip = zip::ZipWriter::new(file);
    let options = FileOptions::<()>::default();

    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.trim().as_bytes()).unwrap();
    zip.add_directory("_rels/", options).unwrap();
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(rels.trim().as_bytes()).unwrap();
    zip.add_directory("word/", options).unwrap();
    zip.add_directory("word/_rels/", options).unwrap();
    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    zip.write_all(rels_xml.trim().as_bytes()).unwrap();
    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(body_xml.trim().as_bytes()).unwrap();

    for (path, xml) in extra_files {
        zip.start_file(*path, options).unwrap();
        zip.write_all(xml.trim().as_bytes()).unwrap();
    }

    zip.finish().unwrap();
    path
}

pub(super) fn build_odf_zip_custom(
    mimetype: &str,
    content_xml: &str,
    manifest_xml: &str,
    extra_files: &[(&str, &[u8])],
) -> Vec<u8> {
    let mut buffer = Vec::new();
    let cursor = std::io::Cursor::new(&mut buffer);
    let mut zip = zip::ZipWriter::new(cursor);
    let stored = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(mimetype.as_bytes()).unwrap();

    zip.start_file("META-INF/manifest.xml", FileOptions::<()>::default())
        .unwrap();
    zip.write_all(manifest_xml.as_bytes()).unwrap();

    zip.start_file("content.xml", FileOptions::<()>::default())
        .unwrap();
    zip.write_all(content_xml.as_bytes()).unwrap();

    zip.start_file("meta.xml", FileOptions::<()>::default())
        .unwrap();
    zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
  <office:meta>
    <dc:title>Parity</dc:title>
  </office:meta>
</office:document-meta>
"#,
        )
        .unwrap();

    for (path, bytes) in extra_files {
        zip.start_file(*path, FileOptions::<()>::default()).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap();
    buffer
}

pub(super) fn create_pptx_with_media(slide_xml: &str, slide_rels: &str) -> std::path::PathBuf {
    let mut path = std::env::temp_dir();
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    path.push(format!("docir_pptx_{nanos}.pptx"));

    let file = File::create(&path).expect("create temp pptx");
    let mut zip = zip::ZipWriter::new(file);
    let options = FileOptions::<()>::default();

    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="wav" ContentType="audio/wav"/>
          <Override PartName="/ppt/presentation.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
          <Override PartName="/ppt/slides/slide1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
        </Types>"#;

    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="ppt/presentation.xml"/>
        </Relationships>"#;

    let presentation = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId r:id="rId1"/>
          </p:sldIdLst>
        </p:presentation>"#;

    let pres_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
            Target="slides/slide1.xml"/>
        </Relationships>"#;

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.trim().as_bytes()).unwrap();
    zip.add_directory("_rels/", options).unwrap();
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(rels.trim().as_bytes()).unwrap();
    zip.add_directory("ppt/", options).unwrap();
    zip.add_directory("ppt/_rels/", options).unwrap();
    zip.start_file("ppt/presentation.xml", options).unwrap();
    zip.write_all(presentation.trim().as_bytes()).unwrap();
    zip.start_file("ppt/_rels/presentation.xml.rels", options)
        .unwrap();
    zip.write_all(pres_rels.trim().as_bytes()).unwrap();
    zip.add_directory("ppt/slides/", options).unwrap();
    zip.add_directory("ppt/slides/_rels/", options).unwrap();
    zip.start_file("ppt/slides/slide1.xml", options).unwrap();
    zip.write_all(slide_xml.trim().as_bytes()).unwrap();
    zip.start_file("ppt/slides/_rels/slide1.xml.rels", options)
        .unwrap();
    zip.write_all(slide_rels.trim().as_bytes()).unwrap();
    zip.add_directory("ppt/media/", options).unwrap();
    zip.start_file("ppt/media/audio1.wav", options).unwrap();
    zip.write_all(b"RIFFDATA").unwrap();
    zip.finish().unwrap();

    path
}
