use super::super::*;
use super::helpers::{build_odf_zip_custom, create_docx_with_relationships};
use docir_core::types::NodeType;
use docir_core::IrSummary;

const DOCX_PARITY_BODY: &str = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
            <w:tbl>
              <w:tr>
                <w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>
              </w:tr>
            </w:tbl>
            <w:p>
              <w:hyperlink r:id="rId3"><w:r><w:t>Link</w:t></w:r></w:hyperlink>
            </w:p>
            <w:p>
              <w:commentRangeStart w:id="0"/>
              <w:r><w:t>Commented</w:t></w:r>
              <w:commentRangeEnd w:id="0"/>
            </w:p>
            <w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>
            <w:p>
              <w:ins w:author="A" w:date="2024-01-01T00:00:00Z">
                <w:r><w:t>Inserted</w:t></w:r>
              </w:ins>
            </w:p>
          </w:body>
        </w:document>
        "#;

const DOCX_PARITY_RELS: &str = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId3"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            Target="https://example.com" TargetMode="External"/>
          <Relationship Id="rId4"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            Target="comments.xml"/>
          <Relationship Id="rId5"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
            Target="footnotes.xml"/>
        </Relationships>
        "#;

const DOCX_PARITY_CONTENT_TYPES: &str = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/comments.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
          <Override PartName="/word/footnotes.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
        </Types>
        "#;

const DOCX_PARITY_COMMENTS: &str = r#"
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:comment w:id="0" w:author="Bob" w:date="2024-01-01T00:00:00Z">
            <w:p><w:r><w:t>Comment text</w:t></w:r></w:p>
          </w:comment>
        </w:comments>
        "#;

const DOCX_PARITY_FOOTNOTES: &str = r#"
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:footnote w:id="1">
            <w:p><w:r><w:t>Footnote</w:t></w:r></w:p>
          </w:footnote>
        </w:footnotes>
        "#;

fn build_docx_parity_fixture() -> std::path::PathBuf {
    create_docx_with_relationships(
        DOCX_PARITY_BODY,
        DOCX_PARITY_RELS,
        DOCX_PARITY_CONTENT_TYPES,
        &[
            ("word/comments.xml", DOCX_PARITY_COMMENTS),
            ("word/footnotes.xml", DOCX_PARITY_FOOTNOTES),
        ],
    )
}

fn build_odt_parity_fixture() -> Vec<u8> {
    let odt_content = r#"
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <text:p>Hello</text:p>
      <table:table>
        <table:table-row>
          <table:table-cell><text:p>Cell</text:p></table:table-cell>
        </table:table-row>
      </table:table>
      <text:p><text:a xlink:href="https://example.com">Link</text:a></text:p>
      <office:annotation>
        <text:p>Comment text</text:p>
      </office:annotation>
      <text:note text:note-class="footnote">
        <text:note-body><text:p>Footnote</text:p></text:note-body>
      </text:note>
      <text:tracked-changes>
        <text:changed-region>
          <text:insertion>
            <text:p>Inserted</text:p>
          </text:insertion>
        </text:changed-region>
      </text:tracked-changes>
    </office:text>
  </office:body>
</office:document-content>
        "#;

    let odt_manifest = r#"
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
        "#;

    build_odf_zip_custom(
        "application/vnd.oasis.opendocument.text",
        odt_content,
        odt_manifest,
        &[],
    )
}

fn assert_docx_odt_parity(docx_summary: &IrSummary, odt_summary: &IrSummary) {
    assert!(docx_summary.run_texts.contains("Hello"));
    assert!(odt_summary.run_texts.contains("Hello"));
    assert!(docx_summary.run_texts.contains("Cell"));
    assert!(odt_summary.run_texts.contains("Cell"));

    assert_eq!(
        docx_summary.count(NodeType::Table),
        odt_summary.count(NodeType::Table)
    );
    assert_eq!(
        docx_summary.count(NodeType::Comment),
        odt_summary.count(NodeType::Comment)
    );
    assert_eq!(
        docx_summary.count(NodeType::Footnote),
        odt_summary.count(NodeType::Footnote)
    );
    assert_eq!(
        docx_summary.count(NodeType::Revision),
        odt_summary.count(NodeType::Revision)
    );
    assert_eq!(
        docx_summary.count(NodeType::ExternalReference),
        odt_summary.count(NodeType::ExternalReference)
    );
}

#[test]
fn test_parity_docx_odt_core_nodes() {
    let docx_path = build_docx_parity_fixture();
    let odt_zip = build_odt_parity_fixture();

    let parser = DocumentParser::new();
    let docx = parser.parse_file(&docx_path).expect("docx parse");
    let odt = parser
        .parse_reader(std::io::Cursor::new(odt_zip))
        .expect("odt parse");

    let docx_summary = IrSummary::from_store(&docx.store);
    let odt_summary = IrSummary::from_store(&odt.store);
    assert_docx_odt_parity(&docx_summary, &odt_summary);

    std::fs::remove_file(docx_path).ok();
}
