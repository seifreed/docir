use super::*;

#[test]
fn test_odf_formula_dde_and_links() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell table:formula="of:=DDE(&quot;soffice&quot;;&quot;file:///tmp/test.ods&quot;;&quot;A1&quot;)" />
          <table:table-cell table:formula="of:=HYPERLINK(&quot;https://example.com&quot;;&quot;Example&quot;)" />
          <table:table-cell table:formula="of:=WEBSERVICE(&quot;https://example.com/api&quot;)" />
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
    let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#;
    let zip_data =
        build_odf_zip_custom(mimetype, content_xml, None, Some(manifest_xml), Vec::new());
    let parser = DocumentParser::new();
    let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    let doc = parsed.document().unwrap();

    assert!(!doc.security.dde_fields.is_empty());
    assert!(
        doc.security
            .threat_indicators
            .iter()
            .any(|i| i.indicator_type == ThreatIndicatorType::DdeCommand)
    );

    let mut has_formula_link = false;
    let mut has_unsupported = false;
    for node in parsed.store.values() {
        if let IRNode::ExternalReference(ext) = node
            && ext.target.contains("example.com")
        {
            has_formula_link = true;
        }
        if let IRNode::Diagnostics(diag) = node
            && diag
                .entries
                .iter()
                .any(|e| e.code == "ODF_FORMULA_UNSUPPORTED_FUNCTION")
        {
            has_unsupported = true;
        }
    }
    assert!(has_formula_link);
    assert!(has_unsupported);
}

#[test]
fn test_odf_encryption_metadata_diagnostics() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:text><text:p>Encrypted</text:p></office:text></office:body>
</office:document-content>
"#;
    let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml">
    <manifest:encryption-data manifest:checksum-type="SHA1" manifest:checksum="YWJjZA==">
      <manifest:algorithm manifest:algorithm-name="http://www.w3.org/2001/04/xmlenc#aes256-cbc"
        manifest:initialisation-vector="MTIzNDU2Nzg5MA==" manifest:key-size="32"/>
      <manifest:key-derivation manifest:key-derivation-name="PBKDF2"
        manifest:salt="c2FsdA==" manifest:iteration-count="2048"/>
    </manifest:encryption-data>
  </manifest:file-entry>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#;
    let zip_data =
        build_odf_zip_custom(mimetype, content_xml, None, Some(manifest_xml), Vec::new());
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

    let mut has_meta = false;
    for node in parsed.store.values() {
        if let IRNode::Diagnostics(diag) = node
            && diag.entries.iter().any(|e| e.code == "ODF_ENCRYPTION_META")
        {
            has_meta = true;
        }
    }
    assert!(has_meta);
}

#[test]
fn test_odf_manifest_inventory_and_parts() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
  <office:body>
    <office:text>
      <text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">Hello</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Thumbnails/thumbnail.png" manifest:media-type="image/png"/>
</manifest:manifest>
"#;
    let zip_data = build_odf_zip_custom(
            mimetype,
            content_xml,
            Some(r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#),
            Some(manifest_xml),
            vec![
                (
                    "settings.xml",
                    br#"<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
                ),
                ("Thumbnails/thumbnail.png", b"pngdata"),
            ],
        );
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

    let mut part_paths = Vec::new();
    let mut asset_paths = Vec::new();
    let mut odf_parts = Vec::new();
    for node in parsed.store.values() {
        match node {
            IRNode::ExtensionPart(part) => part_paths.push(part.path.clone()),
            IRNode::MediaAsset(asset) => asset_paths.push(asset.path.clone()),
            IRNode::Diagnostics(diag) => {
                for entry in &diag.entries {
                    if entry.code == "ODF_PART"
                        && let Some(path) = entry.path.as_ref()
                    {
                        odf_parts.push(path.clone());
                    }
                }
            }
            _ => {}
        }
    }

    assert!(part_paths.contains(&"content.xml".to_string()));
    assert!(part_paths.contains(&"styles.xml".to_string()));
    assert!(part_paths.contains(&"settings.xml".to_string()));
    assert!(asset_paths.contains(&"Thumbnails/thumbnail.png".to_string()));
    assert!(odf_parts.contains(&"content.xml".to_string()));
    assert!(odf_parts.contains(&"styles.xml".to_string()));
}

#[test]
fn test_odt_headers_and_footers_from_styles() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:p>Body</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:master-styles>
    <style:master-page style:name="Standard">
      <style:header>
        <text:p>Header text</text:p>
      </style:header>
      <style:footer>
        <text:p>Footer text</text:p>
      </style:footer>
    </style:master-page>
  </office:master-styles>
</office:document-styles>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, Some(styles_xml));
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

    let mut header_texts = Vec::new();
    let mut footer_texts = Vec::new();
    for node in parsed.store.values() {
        match node {
            IRNode::Header(header) => {
                for id in &header.content {
                    if let Some(IRNode::Paragraph(p)) = parsed.store.get(*id) {
                        for run_id in &p.runs {
                            if let Some(IRNode::Run(run)) = parsed.store.get(*run_id) {
                                header_texts.push(run.text.clone());
                            }
                        }
                    }
                }
            }
            IRNode::Footer(footer) => {
                for id in &footer.content {
                    if let Some(IRNode::Paragraph(p)) = parsed.store.get(*id) {
                        for run_id in &p.runs {
                            if let Some(IRNode::Run(run)) = parsed.store.get(*run_id) {
                                footer_texts.push(run.text.clone());
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    assert!(header_texts.iter().any(|t| t == "Header text"));
    assert!(footer_texts.iter().any(|t| t == "Footer text"));
}

#[test]
fn test_parse_ods_named_ranges_and_pivots() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:spreadsheet>
      <table:named-expressions>
        <table:named-range table:name="RANGE1" table:cell-range-address="Sheet1.A1:Sheet1.B2"/>
        <table:named-expression table:name="EXPR1" table:expression="of:=SUM([.A1];[.B1])"/>
      </table:named-expressions>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell table:cell-value-type="float" table:cell-value="1"/>
          <table:table-cell table:cell-value-type="float" table:cell-value="2"/>
        </table:table-row>
      </table:table>
      <table:data-pilot-table table:name="Pivot1"
        table:source-range-address="Sheet1.A1:Sheet1.B2"
        table:target-range-address="Sheet1.D1:Sheet1.E2">
        <table:data-pilot-field table:source-field-name="Field1"/>
      </table:data-pilot-table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();
    let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    let doc = parsed.document().unwrap();

    assert!(!doc.defined_names.is_empty());

    let mut pivot_tables = 0;
    let mut pivot_caches = 0;
    let mut pivot_records = 0;
    for node in parsed.store.values() {
        match node {
            IRNode::PivotTable(_) => pivot_tables += 1,
            IRNode::PivotCache(_) => pivot_caches += 1,
            IRNode::PivotCacheRecords(_) => pivot_records += 1,
            _ => {}
        }
    }

    assert!(pivot_tables >= 1);
    assert!(pivot_caches >= 1);
    assert!(pivot_records >= 1);
}

#[test]
fn test_parse_ods_named_ranges_validations_and_pivots_accept_alternate_namespace_prefixes() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<pkg:document-content xmlns:pkg="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:calc="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:txt="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <pkg:body>
    <calc:spreadsheet>
      <t:named-expressions>
        <t:named-range t:name="RANGE_ALT" t:cell-range-address="Sheet1.A1:Sheet1.B2"/>
        <t:named-expression t:name="EXPR_ALT" t:expression="of:=SUM([.A1];[.B1])"/>
      </t:named-expressions>
      <t:content-validations>
        <t:content-validation t:name="rule-alt" t:condition="cell-content-is-between(1,10)" />
      </t:content-validations>
      <t:table t:name="Sheet1">
        <t:table-row>
          <t:table-cell t:cell-value-type="float" t:cell-value="1" t:content-validation-name="rule-alt"/>
          <t:table-cell t:cell-value-type="float" t:cell-value="2"/>
        </t:table-row>
      </t:table>
      <t:data-pilot-table t:name="PivotAlt"
        t:source-range-address="Sheet1.A1:Sheet1.B2"
        t:target-range-address="Sheet1.D1:Sheet1.E2">
        <t:data-pilot-field t:source-field-name="Field1"/>
      </t:data-pilot-table>
    </calc:spreadsheet>
  </pkg:body>
</pkg:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();
    let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    let doc = parsed.document().unwrap();

    assert!(doc.defined_names.len() >= 2);

    let mut data_validations = 0;
    let mut pivot_tables = 0;
    let mut pivot_caches = 0;
    let mut pivot_records = 0;
    let mut pivot_named = false;
    for node in parsed.store.values() {
        match node {
            IRNode::DataValidation(_) => data_validations += 1,
            IRNode::PivotTable(pivot) => {
                pivot_tables += 1;
                if pivot.name.as_deref() == Some("PivotAlt")
                    && pivot.ref_range.as_deref() == Some("Sheet1.D1:Sheet1.E2")
                {
                    pivot_named = true;
                }
            }
            IRNode::PivotCache(_) => pivot_caches += 1,
            IRNode::PivotCacheRecords(_) => pivot_records += 1,
            _ => {}
        }
    }

    assert_eq!(data_validations, 1);
    assert_eq!(pivot_tables, 1);
    assert_eq!(pivot_caches, 1);
    assert_eq!(pivot_records, 1);
    assert!(pivot_named);
}

#[test]
fn test_parse_ods_reports_malformed_pivot_start_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:data-pilot-table table:name="Pivot1" table:name="Pivot2"
        table:source-range-address="Sheet1.A1:Sheet1.B2">
        <table:data-pilot-field table:source-field-name="Field1"/>
      </table:data-pilot-table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed pivot attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_ods_reports_malformed_empty_pivot_attributes() {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:spreadsheet>
      <table:data-pilot-table table:name="Pivot1" table:name="Pivot2"
        table:source-range-address="Sheet1.A1:Sheet1.B2"/>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, None);
    let parser = DocumentParser::new();

    let err = parser
        .parse_reader(Cursor::new(zip_data))
        .expect_err("malformed empty pivot attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_odp_master_pages_and_transitions() {
    let mimetype = "application/vnd.oasis.opendocument.presentation";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1" presentation:transition-type="fade" presentation:transition-speed="fast"/>
    </office:presentation>
  </office:body>
</office:document-content>
"#;
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
  <office:master-styles>
    <style:master-page style:name="Master1"/>
  </office:master-styles>
</office:document-styles>
"#;
    let zip_data = build_odf_zip(mimetype, content_xml, Some(styles_xml));
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

    let mut slide_with_transition = 0;
    let mut master_page_diag = false;
    for node in parsed.store.values() {
        if let IRNode::Slide(slide) = node
            && slide.transition.is_some()
        {
            slide_with_transition += 1;
        }
        if let IRNode::Diagnostics(diag) = node
            && diag.entries.iter().any(|e| e.code == "ODF_MASTER_PAGE")
        {
            master_page_diag = true;
        }
    }

    assert_eq!(slide_with_transition, 1);
    assert!(master_page_diag);
}
