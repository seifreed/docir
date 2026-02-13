use crate::xml_utils::local_name;
use docir_core::ir::{ChartData, ChartSeries, IRNode};
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::Event;
use quick_xml::Reader;

#[derive(Clone, Copy)]
enum SeriesSection {
    Name,
    Category,
    Value,
}

pub fn parse_chart_data(xml: &str, chart_path: &str, store: &mut IrStore) -> Option<NodeId> {
    let mut chart = ChartData::new();
    chart.span = Some(SourceSpan::new(chart_path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut in_title = false;
    let mut in_series = false;
    let mut section: Option<SeriesSection> = None;
    let mut current_series: Option<ChartSeries> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let name = local_name(&name_buf);
                handle_start_event(
                    name,
                    &mut chart,
                    &mut in_series,
                    &mut in_title,
                    &mut section,
                    &mut current_series,
                );
            }
            Ok(Event::Text(e)) => {
                let text = e.unescape().unwrap_or_default().to_string();
                handle_text_event(
                    &text,
                    &mut chart,
                    in_series,
                    in_title,
                    section,
                    current_series.as_mut(),
                );
            }
            Ok(Event::End(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let name = local_name(&name_buf);
                handle_end_event(
                    name,
                    &mut chart,
                    &mut in_series,
                    &mut in_title,
                    &mut section,
                    &mut current_series,
                );
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

fn handle_start_event<'a>(
    name: &'a [u8],
    chart: &mut ChartData,
    in_series: &mut bool,
    in_title: &mut bool,
    section: &mut Option<SeriesSection>,
    current_series: &mut Option<ChartSeries>,
) {
    if name.ends_with(b"Chart") {
        chart.chart_type = Some(String::from_utf8_lossy(name).to_string());
    }
    if name == b"ser" {
        *in_series = true;
        *current_series = Some(ChartSeries::new());
    }
    if !*in_series && name == b"title" {
        *in_title = true;
    }
    if *in_series {
        *section = match name {
            b"tx" => Some(SeriesSection::Name),
            b"cat" => Some(SeriesSection::Category),
            b"val" => Some(SeriesSection::Value),
            _ => *section,
        };
    }
}

fn handle_text_event(
    text: &str,
    chart: &mut ChartData,
    in_series: bool,
    in_title: bool,
    section: Option<SeriesSection>,
    current_series: Option<&mut ChartSeries>,
) {
    if in_title && chart.title.is_none() {
        chart.title = Some(text.to_string());
        return;
    }
    if !in_series {
        return;
    }
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return;
    }
    let Some(series) = current_series else {
        return;
    };
    match section {
        Some(SeriesSection::Name) if series.name.is_none() => {
            series.name = Some(trimmed.to_string());
        }
        Some(SeriesSection::Category) => {
            series.categories.push(trimmed.to_string());
        }
        Some(SeriesSection::Value) => {
            series.values.push(trimmed.to_string());
        }
        _ => {}
    }
}

fn handle_end_event(
    name: &[u8],
    chart: &mut ChartData,
    in_series: &mut bool,
    in_title: &mut bool,
    section: &mut Option<SeriesSection>,
    current_series: &mut Option<ChartSeries>,
) {
    if name == b"title" {
        *in_title = false;
    }
    if name == b"ser" {
        *in_series = false;
        *section = None;
        if let Some(series) = current_series.take() {
            if let Some(name) = &series.name {
                chart.series.push(name.clone());
            }
            chart.series_data.push(series);
        }
    }
    if matches!(name, b"tx" | b"cat" | b"val") {
        *section = None;
    }
}
