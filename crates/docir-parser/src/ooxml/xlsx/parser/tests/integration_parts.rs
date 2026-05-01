use super::*;
use crate::ooxml::xlsx::{
    parse_external_link_part, parse_query_table_part, parse_slicer_part, parse_timeline_part,
};
use docir_core::ir::IRNode;

#[test]
fn test_parse_people_part() {
    let xml = r#"
        <ppl:people xmlns:ppl="http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes">
          <ppl:person ppl:id="p1" ppl:userId="user1" ppl:displayName="Alice" ppl:initials="A"/>
          <ppl:person ppl:id="p2" ppl:userId="user2" ppl:displayName="Bob"/>
        </ppl:people>
        "#;
    let mut parser = XlsxParser::new();
    let mut zip = build_zip_with_entries(vec![("xl/persons/person.xml", xml)]);
    let workbook_xml =
        r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
    let doc_id = parser
        .parse_workbook(
            &mut zip,
            workbook_xml,
            &Relationships::default(),
            "xl/workbook.xml",
        )
        .expect("workbook");
    let store = parser.into_store();

    let doc = match store.get(doc_id) {
        Some(IRNode::Document(d)) => d,
        _ => panic!("missing document"),
    };
    assert_eq!(doc.shared_parts.len(), 1);

    let people = match store.get(doc.shared_parts[0]) {
        Some(IRNode::PeoplePart(p)) => p,
        _ => panic!("missing people part"),
    };
    assert_eq!(people.people.len(), 2);
    assert_eq!(people.people[0].display_name.as_deref(), Some("Alice"));
    assert_eq!(people.people[1].display_name.as_deref(), Some("Bob"));
}

#[test]
fn test_parse_external_link_part() {
    let xml = r#"
        <externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <externalBook r:id="rId1">
            <sheetNames>
              <sheetName val="SheetA"/>
              <sheetName val="SheetB"/>
            </sheetNames>
          </externalBook>
        </externalLink>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
            Target="file:///C:/data.xlsx"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels");
    let part = parse_external_link_part(xml, "xl/externalLinks/externalLink1.xml", Some(&rels))
        .expect("external link");
    assert_eq!(part.target.as_deref(), Some("file:///C:/data.xlsx"));
    assert_eq!(part.sheets.len(), 2);
    assert_eq!(part.sheets[0].name.as_deref(), Some("SheetA"));
}

#[test]
fn test_parse_slicer_part() {
    let xml = r#"
        <slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                name="Slicer1" caption="Region" cache="1" />
        "#;
    let slicer = parse_slicer_part(xml, "xl/slicers/slicer1.xml").expect("slicer");
    assert_eq!(slicer.name.as_deref(), Some("Slicer1"));
    assert_eq!(slicer.caption.as_deref(), Some("Region"));
    assert_eq!(slicer.cache_id.as_deref(), Some("1"));
}

#[test]
fn test_parse_timeline_part() {
    let xml = r#"
        <timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                  name="Timeline1" cache="2" />
        "#;
    let timeline = parse_timeline_part(xml, "xl/timelines/timeline1.xml").expect("timeline");
    assert_eq!(timeline.name.as_deref(), Some("Timeline1"));
    assert_eq!(timeline.cache_id.as_deref(), Some("2"));
}

#[test]
fn test_parse_query_table_part() {
    let xml = r#"
        <queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    name="Query1" connectionId="7">
          <dbPr command="SELECT * FROM tbl"/>
          <webPr url="https://example.com/data"/>
        </queryTable>
        "#;
    let query = parse_query_table_part(xml, "xl/queryTables/queryTable1.xml").expect("query");
    assert_eq!(query.name.as_deref(), Some("Query1"));
    assert_eq!(query.connection_id.as_deref(), Some("7"));
    assert_eq!(query.command.as_deref(), Some("SELECT * FROM tbl"));
    assert_eq!(query.url.as_deref(), Some("https://example.com/data"));
}

#[test]
fn test_parse_workbook_loads_optional_parts_bundle() {
    let workbook_xml = r#"
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <bookViews>
            <workbookView activeTab="0" firstSheet="0"/>
          </bookViews>
          <definedNames>
            <definedName name="_xlnm.Auto_Open">Macro1!$A$1</definedName>
          </definedNames>
        </workbook>
        "#;
    let shared_strings_xml = r#"
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <si><t>Hello</t></si>
        </sst>
        "#;
    let styles_xml = r#"
        <styleSheet>
          <fonts count="1"><font><name val="Calibri"/></font></fonts>
          <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
          <borders count="1"><border/></borders>
          <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
        </styleSheet>
        "#;
    let calc_chain_xml = r#"
        <calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <c r="A1"/>
        </calcChain>
        "#;
    let people_xml = r#"
        <ppl:people xmlns:ppl="http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes">
          <ppl:person ppl:id="p1" ppl:userId="u1" ppl:displayName="Alice"/>
        </ppl:people>
        "#;
    let metadata_xml = r#"
        <metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <metadataTypes count="1">
            <metadataType name="XLRICHVALUE" copy="1" update="1"/>
          </metadataTypes>
          <cellMetadata count="1"/>
          <valueMetadata count="0"/>
        </metadata>
        "#;
    let slicer_xml = r#"
        <slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                name="Slicer1" caption="Region" cache="1"/>
        "#;
    let timeline_xml = r#"
        <timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                  name="Timeline1" cache="2"/>
        "#;
    let query_xml = r#"
        <queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    name="Query1" connectionId="7">
          <dbPr command="SELECT * FROM t"/>
        </queryTable>
        "#;

    let mut zip = build_zip_with_entries(vec![
        ("xl/sharedStrings.xml", shared_strings_xml),
        ("xl/styles.xml", styles_xml),
        ("xl/calcChain.xml", calc_chain_xml),
        ("xl/persons/person.xml", people_xml),
        ("xl/metadata.xml", metadata_xml),
        ("xl/slicers/slicer1.xml", slicer_xml),
        ("xl/timelines/timeline1.xml", timeline_xml),
        ("xl/queryTables/queryTable1.xml", query_xml),
    ]);

    let mut parser = XlsxParser::new();
    let doc_id = parser
        .parse_workbook(
            &mut zip,
            workbook_xml,
            &Relationships::default(),
            "xl/workbook.xml",
        )
        .expect("workbook");
    let store = parser.into_store();
    let doc = match store.get(doc_id) {
        Some(IRNode::Document(d)) => d,
        _ => panic!("missing document"),
    };

    assert!(doc.workbook_properties.is_some());
    assert!(doc.shared_strings.is_some());
    assert!(doc.spreadsheet_styles.is_some());
    assert!(doc.sheet_metadata.is_some());
    assert!(doc
        .shared_parts
        .iter()
        .any(|id| matches!(store.get(*id), Some(IRNode::CalcChain(_)))));
    assert!(doc
        .shared_parts
        .iter()
        .any(|id| matches!(store.get(*id), Some(IRNode::PeoplePart(_)))));
    assert!(doc
        .shared_parts
        .iter()
        .any(|id| matches!(store.get(*id), Some(IRNode::SlicerPart(_)))));
    assert!(doc
        .shared_parts
        .iter()
        .any(|id| matches!(store.get(*id), Some(IRNode::TimelinePart(_)))));
    assert!(doc
        .shared_parts
        .iter()
        .any(|id| matches!(store.get(*id), Some(IRNode::QueryTablePart(_)))));
    assert!(!doc.defined_names.is_empty());
}

#[test]
fn test_load_worksheet_comments_collects_legacy_and_threaded() {
    let mut parser = XlsxParser::new();
    let mut zip = build_zip_with_entries(vec![
        (
            "xl/comments1.xml",
            r#"
            <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <authors><author>Alice</author></authors>
              <commentList>
                <comment ref="A1" authorId="0"><text><t>Hello</t></text></comment>
              </commentList>
            </comments>
            "#,
        ),
        (
            "xl/threadedComments/threadedComment1.xml",
            r#"
            <ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
              <threadedComment ref="B2" personId="person-1"><text>Thread</text></threadedComment>
            </ThreadedComments>
            "#,
        ),
    ]);
    let rels = Relationships::parse(
        r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdCom"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            Target="../comments1.xml"/>
          <Relationship Id="rIdThr"
            Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
            Target="../threadedComments/threadedComment1.xml"/>
        </Relationships>
        "#,
    )
    .expect("rels");

    let ids = parser
        .load_worksheet_comments(&mut zip, "xl/worksheets/sheet1.xml", &rels, "Sheet1")
        .expect("comments");
    assert_eq!(ids.len(), 2);

    let store = parser.into_store();
    let legacy = match store.get(ids[0]) {
        Some(IRNode::SheetComment(c)) => c,
        _ => panic!("expected legacy comment"),
    };
    let threaded = match store.get(ids[1]) {
        Some(IRNode::SheetComment(c)) => c,
        _ => panic!("expected threaded comment"),
    };
    assert_eq!(legacy.cell_ref, "A1");
    assert_eq!(legacy.author.as_deref(), Some("Alice"));
    assert_eq!(threaded.cell_ref, "B2");
    assert_eq!(threaded.author.as_deref(), Some("person-1"));
}

#[test]
fn test_load_worksheet_tables_and_pivots_skips_missing_targets() {
    let mut parser = XlsxParser::new();
    let mut zip = build_zip_with_entries(vec![
        (
            "xl/tables/table1.xml",
            r#"<table name="Table1" displayName="Table1" ref="A1:B2"/>"#,
        ),
        (
            "xl/pivotTables/pivotTable1.xml",
            r#"<pivotTableDefinition name="Pivot1" cacheId="3"><location ref="D1:F10"/></pivotTableDefinition>"#,
        ),
    ]);
    let rels = Relationships::parse(
        r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdTbl1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
            Target="../tables/table1.xml"/>
          <Relationship Id="rIdTblMissing"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
            Target="../tables/missing.xml"/>
          <Relationship Id="rIdPivot1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
            Target="../pivotTables/pivotTable1.xml"/>
        </Relationships>
        "#,
    )
    .expect("rels");

    let table_ids = parser
        .load_worksheet_tables(&mut zip, "xl/worksheets/sheet1.xml", &rels)
        .expect("tables");
    let pivot_ids = parser
        .load_worksheet_pivots(&mut zip, "xl/worksheets/sheet1.xml", &rels)
        .expect("pivots");

    assert_eq!(table_ids.len(), 1);
    assert_eq!(pivot_ids.len(), 1);

    let store = parser.into_store();
    assert!(matches!(
        store.get(table_ids[0]),
        Some(IRNode::TableDefinition(_))
    ));
    assert!(matches!(
        store.get(pivot_ids[0]),
        Some(IRNode::PivotTable(_))
    ));
}
