use crate::xml_utils::local_name;
use docir_core::ir::{ChartData, ChartSeries, IRNode};
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::Event;
use quick_xml::Reader;

pub fn parse_chart_data(xml: &str, chart_path: &str, store: &mut IrStore) -> Option<NodeId> {
    let mut chart = ChartData::new();
    chart.span = Some(SourceSpan::new(chart_path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut in_title = false;
    let mut in_series = false;
    let mut section: Option<&[u8]> = None;
    let mut current_series: Option<ChartSeries> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let name = local_name(&name_buf);
                if name.ends_with(b"Chart") {
                    chart.chart_type = Some(String::from_utf8_lossy(name).to_string());
                }
                if name == b"ser" {
                    in_series = true;
                    current_series = Some(ChartSeries::new());
                }
                if !in_series && name == b"title" {
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
            Ok(Event::End(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let name = local_name(&name_buf);
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
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    let id = chart.id;
    store.insert(IRNode::ChartData(chart));
    Some(id)
}
