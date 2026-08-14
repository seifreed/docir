use super::{
    scan_embedded_objects, scan_external_links, scan_odf_filters, scan_odf_formula_security,
    scan_odf_objects,
};
use crate::error::ParseError;
use crate::zip_handler::PackageReader;
use docir_core::security::ExternalRefType;

struct FailingSizeReader;

impl PackageReader for FailingSizeReader {
    fn contains(&self, _name: &str) -> bool {
        false
    }

    fn read_file(&mut self, name: &str) -> Result<Vec<u8>, ParseError> {
        Err(ParseError::MissingPart(name.to_string()))
    }

    fn file_size(&mut self, name: &str) -> Result<u64, ParseError> {
        Err(ParseError::MissingPart(name.to_string()))
    }

    fn file_names(&self) -> Vec<String> {
        Vec::new()
    }

    fn list_prefix(&self, _prefix: &str) -> Vec<String> {
        Vec::new()
    }

    fn list_suffix(&self, _suffix: &str) -> Vec<String> {
        Vec::new()
    }
}

#[test]
fn scan_external_links_accepts_alternate_namespace_prefixes() {
    let xml = r#"
            <office:document-content
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:txt="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
              xmlns:dr="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
              xmlns:lnk="http://www.w3.org/1999/xlink">
              <txt:a lnk:href="https://example.test/link">Link</txt:a>
              <dr:image lnk:href="Pictures/pic.png"/>
              <dr:object-ole lnk:href="https://example.test/object.bin"/>
            </office:document-content>
        "#;

    let refs = scan_external_links(xml, "content.xml").expect("external links");
    let types: Vec<_> = refs.iter().map(|r| r.ref_type).collect();

    assert_eq!(refs.len(), 2);
    assert!(types.contains(&ExternalRefType::Hyperlink));
    assert!(types.contains(&ExternalRefType::Image));
}

#[test]
fn scan_external_links_reports_invalid_attribute_entity() {
    let xml = r#"<office:document-content
        xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:lnk="http://www.w3.org/1999/xlink">
        <office:a lnk:href="https://example.test/bad &"/>
    </office:document-content>"#;

    match scan_external_links(xml, "styles.xml")
        .expect_err("invalid external link entity must fail")
    {
        ParseError::Xml { file, .. } => assert_eq!(file, "styles.xml"),
        other => panic!("expected XML error, got {other:?}"),
    }
}

#[test]
fn scan_odf_objects_accepts_alternate_namespace_prefixes() {
    let xml = r#"
            <office:document-content
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:dr="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
              xmlns:lnk="http://www.w3.org/1999/xlink">
              <dr:object-ole lnk:href="https://example.test/object.bin"/>
            </office:document-content>
        "#;

    let (oles, refs) = scan_odf_objects(xml).expect("odf objects");
    assert_eq!(oles.len(), 1);
    assert_eq!(
        oles[0].link_target.as_deref(),
        Some("https://example.test/object.bin")
    );
    assert_eq!(refs.len(), 1);
    assert_eq!(refs[0].ref_type, ExternalRefType::OleLink);
}

#[test]
fn scan_odf_object_internal_targets_are_not_external_ole_links() {
    let xml = r#"
            <office:document-content
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:dr="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
              xmlns:lnk="http://www.w3.org/1999/xlink">
              <dr:object-ole lnk:href="Object 1"/>
            </office:document-content>
        "#;

    let refs = scan_external_links(xml, "content.xml").expect("external links");
    let (oles, ole_refs) = scan_odf_objects(xml).expect("odf objects");

    assert!(refs.is_empty());
    assert_eq!(oles.len(), 1);
    assert!(!oles[0].is_linked);
    assert!(ole_refs.is_empty());
}

#[test]
fn scan_embedded_objects_reports_size_errors() {
    let mut reader = FailingSizeReader;
    let err = scan_embedded_objects(&["Object 1".to_string()], &mut reader)
        .expect_err("embedded object size failures must not be suppressed");

    match err {
        ParseError::MissingPart(path) => assert_eq!(path, "Object 1"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn scan_odf_formula_security_accepts_alternate_formula_prefix() {
    let xml = r#"
            <office:document-content
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:tbl="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
              <tbl:table-cell tbl:formula='of:=DDE("cmd";"/c calc";"A1")'/>
            </office:document-content>
        "#;

    let scan = scan_odf_formula_security(xml).expect("formula security");
    assert_eq!(scan.dde_fields.len(), 1);
    assert_eq!(scan.dde_fields[0].application, "cmd");
}

#[test]
fn scan_odf_filters_accepts_alternate_namespace_prefixes() {
    let xml = r#"
            <office:document-content
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:tbl="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
              <tbl:filter tbl:target-range-address="Sheet1.A1:Sheet1.B2"/>
              <tbl:filter-and tbl:condition="cell-content()&gt;0"/>
            </office:document-content>
        "#;

    assert_eq!(
        scan_odf_filters(xml).expect("odf filters"),
        vec![
            "Sheet1.A1:Sheet1.B2".to_string(),
            "cell-content()>0".to_string()
        ]
    );
}

#[test]
fn scan_odf_security_helpers_report_malformed_xml() {
    let malformed = r#"
            <office:document-content
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:tbl="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
              <tbl:table-cell tbl:formula='of:=SUM([.A1])'
        "#;

    let formula_err = match scan_odf_formula_security(malformed) {
        Ok(_) => panic!("formula scan must fail"),
        Err(err) => err,
    };

    for err in [
        scan_external_links(malformed, "content.xml").expect_err("external links must fail"),
        formula_err,
        scan_odf_filters(malformed).expect_err("filter scan must fail"),
    ] {
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
            other => panic!("expected content.xml parse error, got {:?}", other),
        }
    }
}

#[test]
fn scan_odf_security_helpers_reject_multiple_roots() {
    let xml = "<document-content/><document-content/>";

    for result in [
        scan_external_links(xml, "content.xml").map(|_| ()),
        scan_odf_objects(xml).map(|_| ()),
        scan_odf_filters(xml).map(|_| ()),
        scan_odf_formula_security(xml).map(|_| ()),
    ] {
        let err = result.expect_err("ODF security XML must have one root");
        assert!(format!("{err}").contains("multiple roots"));
    }
}
