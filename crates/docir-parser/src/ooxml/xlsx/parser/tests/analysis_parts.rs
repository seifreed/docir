use super::*;
use crate::ooxml::xlsx::{connection_targets, parse_connections_part, parse_sheet_metadata};
use docir_core::ir::IRNode;

#[test]
fn test_parse_chart_xml() {
    let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart>
            <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:lineChart>
              <c:ser><c:tx><c:v>Q1</c:v></c:tx></c:ser>
            </c:lineChart>
          </c:chart>
        </c:chartSpace>
        "#;
    let mut parser = XlsxParser::new();
    let id = parser
        .parse_chart(xml, "xl/charts/chart2.xml")
        .expect("chart");
    let store = parser.into_store();
    let chart = match store.get(id) {
        Some(IRNode::ChartData(c)) => c,
        _ => panic!("missing chart"),
    };
    assert!(chart
        .chart_type
        .as_deref()
        .unwrap_or("")
        .contains("lineChart"));
    assert_eq!(chart.title.as_deref(), Some("Revenue"));
    assert_eq!(chart.series.len(), 1);
    assert_eq!(chart.series_data.len(), 1);
}

#[test]
fn test_parse_connections_xml_targets() {
    let xml = r#"
        <connections>
          <connection id="1" name="Conn1" type="1">
            <webPr url="https://example.com/data"/>
          </connection>
          <connection id="2" name="Conn2" type="2">
            <dbPr connection="DatabaseName" command="SELECT * FROM foo" commandType="2"/>
          </connection>
        </connections>
        "#;
    let part = parse_connections_part(xml, "xl/connections.xml").expect("connections");
    assert_eq!(part.entries.len(), 2);
    assert_eq!(part.entries[0].connection_id, Some(1));
    assert_eq!(
        part.entries[0].url.as_deref(),
        Some("https://example.com/data")
    );
    assert_eq!(part.entries[1].connection_id, Some(2));
    assert_eq!(part.entries[1].connection.as_deref(), Some("DatabaseName"));
    assert_eq!(
        part.entries[1].command.as_deref(),
        Some("SELECT * FROM foo")
    );
    assert_eq!(part.entries[1].command_type, Some(2));
    let targets = connection_targets(&part);
    assert!(targets.contains(&"https://example.com/data".to_string()));
    assert!(targets.contains(&"DatabaseName".to_string()));
}

#[test]
fn test_parse_sheet_metadata() {
    let xml = r#"
        <metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <metadataTypes count="1">
            <metadataType name="XLDynamicArray" minSupportedVersion="120000" copy="1" update="0"/>
          </metadataTypes>
          <cellMetadata count="2"/>
          <valueMetadata count="3"/>
        </metadata>
        "#;
    let meta = parse_sheet_metadata(xml, "xl/metadata.xml").expect("metadata");
    assert_eq!(meta.metadata_types.len(), 1);
    assert_eq!(
        meta.metadata_types[0].name.as_deref(),
        Some("XLDynamicArray")
    );
    assert_eq!(meta.cell_metadata_count, Some(2));
    assert_eq!(meta.value_metadata_count, Some(3));
}
