use super::*;
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
