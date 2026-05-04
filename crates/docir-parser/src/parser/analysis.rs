use crate::error::ParseError;
use crate::xml_utils::local_name;
use crate::xml_utils::lossy_attr_value;
use docir_core::ir::{CellError, IRNode};
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;

pub(super) fn parse_activex_xml(
    xml: &str,
    _path: &str,
) -> Option<docir_core::security::ActiveXControl> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut control = docir_core::security::ActiveXControl::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    let value = lossy_attr_value(&attr).to_string();
                    match key {
                        b"name" => control.name = Some(value.clone()),
                        b"clsid" | b"classid" => control.clsid = Some(value.clone()),
                        b"progid" => control.prog_id = Some(value.clone()),
                        _ => {
                            let prop_key = String::from_utf8_lossy(key).to_string();
                            control.properties.push((prop_key, value));
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    if control.name.is_some() || control.clsid.is_some() || control.prog_id.is_some() {
        Some(control)
    } else {
        None
    }
}

pub(super) fn map_calamine_error(err: calamine::CellErrorType) -> CellError {
    use calamine::CellErrorType::*;
    match err {
        Null => CellError::Null,
        Div0 => CellError::DivZero,
        Value => CellError::Value,
        Ref => CellError::Ref,
        Name => CellError::Name,
        Num => CellError::Num,
        NA => CellError::NA,
        GettingData => CellError::GettingData,
    }
}

pub(super) fn parse_smartart_part(
    xml: &str,
    path: &str,
) -> Result<docir_core::ir::SmartArtPart, ParseError> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut root_element: Option<String> = None;
    let mut point_count: u32 = 0;
    let mut connection_count: u32 = 0;
    let mut rel_ids: Vec<String> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if root_element.is_none() {
                    root_element = Some(String::from_utf8_lossy(e.name().as_ref()).to_string());
                }
                let name_buf = e.name().as_ref().to_vec();
                let name = name_buf.as_slice();
                if name.ends_with(b":pt") || name == b"dgm:pt" {
                    point_count += 1;
                }
                if name.ends_with(b":cxn") || name == b"dgm:cxn" {
                    connection_count += 1;
                }
                if name.ends_with(b":relIds") || name == b"dgm:relIds" {
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        if key == b"r:dm" || key == b"r:lo" || key == b"r:qs" || key == b"r:cs" {
                            let value = lossy_attr_value(&attr).to_string();
                            if !value.is_empty() {
                                rel_ids.push(value);
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(docir_core::ir::SmartArtPart {
        id: NodeId::new(),
        kind: "diagram".to_string(),
        path: path.to_string(),
        root_element,
        point_count: if point_count > 0 {
            Some(point_count)
        } else {
            None
        },
        connection_count: if connection_count > 0 {
            Some(connection_count)
        } else {
            None
        },
        rel_ids,
        span: Some(SourceSpan::new(path)),
    })
}

pub(super) fn parse_chart_data(
    xml: &str,
    chart_path: &str,
    store: &mut IrStore,
) -> Result<NodeId, ParseError> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut chart = docir_core::ir::ChartData::new();
    chart.span = Some(SourceSpan::new(chart_path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_title = false;
    let mut in_series = false;
    let mut section: Option<&[u8]> = None;
    let mut current_series: Option<docir_core::ir::ChartSeries> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name_bytes = e.name().as_ref().to_vec();
                let name = local_name(&name_bytes);
                if name.ends_with(b"Chart") || name.ends_with(b"chart") {
                    chart.chart_type = Some(String::from_utf8_lossy(name).to_string());
                }
                if name == b"ser" {
                    in_series = true;
                    current_series = Some(docir_core::ir::ChartSeries::new());
                }
                if !in_series && (name == b"title") {
                    in_title = true;
                }
                if in_series {
                    if name == b"tx" {
                        section = Some(b"tx");
                    } else if name == b"cat" {
                        section = Some(b"cat");
                    } else if name == b"val" {
                        section = Some(b"val");
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name_bytes = e.name().as_ref().to_vec();
                let name = local_name(&name_bytes);
                if name == b"title" {
                    in_title = false;
                }
                if name == b"ser" {
                    in_series = false;
                    section = None;
                    if let Some(series) = current_series.take() {
                        if let Some(name) = &series.name {
                            chart.series.push(name.clone());
                        }
                        chart.series_data.push(series);
                    }
                }
                if name == b"tx" || name == b"cat" || name == b"val" {
                    section = None;
                }
            }
            Ok(Event::Text(e)) => {
                let text = e.unescape().unwrap_or_default().to_string();
                if in_title && chart.title.is_none() {
                    chart.title = Some(text);
                } else if in_series {
                    let trimmed = text.trim();
                    if trimmed.is_empty() {
                        // skip
                    } else if let Some(series) = current_series.as_mut() {
                        match section {
                            Some(b"tx") => {
                                if series.name.is_none() {
                                    series.name = Some(trimmed.to_string());
                                }
                            }
                            Some(b"cat") => {
                                series.categories.push(trimmed.to_string());
                            }
                            Some(b"val") => {
                                series.values.push(trimmed.to_string());
                            }
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: chart_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    let id = chart.id;
    store.insert(IRNode::ChartData(chart));
    Ok(id)
}

/// Helper function to encode bytes as hex.
pub(super) mod hex {
    /// Public API entrypoint: encode.
    pub fn encode(data: impl AsRef<[u8]>) -> String {
        data.as_ref().iter().map(|b| format!("{:02x}", b)).collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::CellErrorType;
    use docir_core::ir::IRNode;

    #[test]
    fn parse_activex_xml_extracts_core_attributes_and_properties() {
        let xml = r#"<ocx:ocx name="CommandButton1" clsid="{ABC}" progid="Forms.CommandButton.1" custom="x"/>"#;
        let control = parse_activex_xml(xml, "word/activeX/activeX1.xml")
            .expect("expected a parsed ActiveX control");

        assert_eq!(control.name.as_deref(), Some("CommandButton1"));
        assert_eq!(control.clsid.as_deref(), Some("{ABC}"));
        assert_eq!(control.prog_id.as_deref(), Some("Forms.CommandButton.1"));
        assert!(control
            .properties
            .iter()
            .any(|(k, v)| k == "custom" && v == "x"));
    }

    #[test]
    fn parse_activex_xml_returns_none_without_identifying_attributes() {
        let xml = r#"<ocx:ocx custom="x" another="y"/>"#;
        assert!(parse_activex_xml(xml, "word/activeX/activeX1.xml").is_none());
    }

    #[test]
    fn map_calamine_error_maps_all_variants() {
        let cases = [
            (CellErrorType::Null, CellError::Null),
            (CellErrorType::Div0, CellError::DivZero),
            (CellErrorType::Value, CellError::Value),
            (CellErrorType::Ref, CellError::Ref),
            (CellErrorType::Name, CellError::Name),
            (CellErrorType::Num, CellError::Num),
            (CellErrorType::NA, CellError::NA),
            (CellErrorType::GettingData, CellError::GettingData),
        ];

        for (input, expected) in cases {
            assert_eq!(map_calamine_error(input), expected);
        }
    }

    #[test]
    fn parse_smartart_part_extracts_counts_and_relationship_ids() {
        let xml = r#"
            <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dgm:pt modelId="0"/>
              <dgm:pt modelId="1"/>
              <dgm:cxn srcId="0" destId="1"/>
              <dgm:relIds r:dm="rId1" r:lo="rId2"/>
            </dgm:dataModel>
        "#;

        let part = parse_smartart_part(xml, "ppt/diagrams/data1.xml").expect("smartart parsed");
        assert_eq!(part.path, "ppt/diagrams/data1.xml");
        assert_eq!(part.root_element.as_deref(), Some("dgm:dataModel"));
        assert_eq!(part.point_count, Some(2));
        assert_eq!(part.connection_count, Some(1));
        assert_eq!(part.rel_ids, vec!["rId1".to_string(), "rId2".to_string()]);
    }

    #[test]
    fn parse_smartart_part_reports_xml_errors_with_source_path() {
        let err = parse_smartart_part("<dgm:dataModel><dgm:pt></dgm:dataModel>", "bad.xml")
            .expect_err("invalid XML should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "bad.xml"),
            other => panic!("unexpected error variant: {other:?}"),
        }
    }

    #[test]
    fn parse_chart_data_extracts_title_and_series_values() {
        let xml = r#"
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart>
                <c:title><c:tx><c:rich><a:t xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">Quarterly Revenue</a:t></c:rich></c:tx></c:title>
                <c:plotArea>
                  <c:barChart>
                    <c:ser>
                      <c:tx><c:v>North</c:v></c:tx>
                      <c:cat><c:v>Q1</c:v><c:v>Q2</c:v></c:cat>
                      <c:val><c:v>10</c:v><c:v>20</c:v></c:val>
                    </c:ser>
                    <c:ser>
                      <c:tx><c:v>South</c:v></c:tx>
                      <c:cat><c:v>Q1</c:v></c:cat>
                      <c:val><c:v>15</c:v></c:val>
                    </c:ser>
                  </c:barChart>
                </c:plotArea>
              </c:chart>
            </c:chartSpace>
        "#;

        let mut store = IrStore::new();
        let chart_id =
            parse_chart_data(xml, "xl/charts/chart1.xml", &mut store).expect("chart parsed");

        let Some(IRNode::ChartData(chart)) = store.get(chart_id) else {
            panic!("expected chart node");
        };
        assert_eq!(chart.chart_type.as_deref(), Some("barChart"));
        assert_eq!(chart.title.as_deref(), Some("Quarterly Revenue"));
        assert_eq!(chart.series, vec!["North".to_string(), "South".to_string()]);
        assert_eq!(chart.series_data.len(), 2);
        assert_eq!(
            chart.series_data[0].categories,
            vec!["Q1".to_string(), "Q2".to_string()]
        );
        assert_eq!(
            chart.series_data[0].values,
            vec!["10".to_string(), "20".to_string()]
        );
        assert_eq!(chart.series_data[1].categories, vec!["Q1".to_string()]);
    }

    #[test]
    fn parse_chart_data_reports_xml_errors_with_source_path() {
        let mut store = IrStore::new();
        let err = parse_chart_data(
            "<c:chart><c:title></c:chartx>",
            "xl/charts/bad.xml",
            &mut store,
        )
        .expect_err("invalid XML");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/charts/bad.xml"),
            other => panic!("unexpected error variant: {other:?}"),
        }
    }
}
