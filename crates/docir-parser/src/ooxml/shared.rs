//! Shared OOXML parts parsing (themes, relationships, custom XML, media).

use crate::error::ParseError;
use crate::ooxml::relationships::{Relationships, TargetMode};
use docir_core::ir::{
    ChartData, ChartSeries, CustomXmlPart, DigitalSignature, DrawingPart, ExtensionPart,
    ExtensionPartKind, IRNode, MediaType, PeoplePart, PersonEntry, RelationshipEntry,
    RelationshipGraph, RelationshipTargetMode, Shape, ShapeType, Theme, ThemeColor,
    ThemeFontScheme, VmlDrawing, VmlShape, WebExtension, WebExtensionProperty,
    WebExtensionTaskpane,
};
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashSet;

pub fn build_relationship_graph(
    source: &str,
    rels_path: &str,
    rels: &Relationships,
) -> RelationshipGraph {
    let mut graph = RelationshipGraph::new(source);
    // Coverage wants to mark the `.rels` part itself as "seen".
    // We still keep `graph.source` to point at the owning part (e.g. word/document.xml).
    graph.span = Some(SourceSpan::new(rels_path));

    for rel in rels.by_id.values() {
        let target_mode = match rel.target_mode {
            TargetMode::Internal => RelationshipTargetMode::Internal,
            TargetMode::External => RelationshipTargetMode::External,
        };
        graph.relationships.push(RelationshipEntry {
            id: rel.id.clone(),
            rel_type: rel.rel_type.clone(),
            target: rel.target.clone(),
            target_mode,
        });
    }

    graph
}

pub fn parse_theme(xml: &str, path: &str) -> Result<Theme, ParseError> {
    let mut theme = Theme::new();
    theme.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_clr_scheme = false;
    let mut current_color_name: Option<String> = None;
    let mut font_scheme = ThemeFontScheme::default();
    let mut in_major_font = false;
    let mut in_minor_font = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"a:theme" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"name" {
                            theme.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"a:clrScheme" => {
                    in_clr_scheme = true;
                }
                b"a:fontScheme" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"name" {
                            if theme.name.is_none() {
                                theme.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                b"a:majorFont" => {
                    in_major_font = true;
                }
                b"a:minorFont" => {
                    in_minor_font = true;
                }
                b"a:latin" => {
                    let mut typeface = None;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"typeface" {
                            typeface = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                    if let Some(tf) = typeface {
                        if in_major_font {
                            font_scheme.major = Some(tf);
                        } else if in_minor_font {
                            font_scheme.minor = Some(tf);
                        }
                    }
                }
                _ => {
                    if in_clr_scheme {
                        let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                        current_color_name = Some(name);
                    }
                }
            },
            Ok(Event::Empty(e)) => {
                if in_clr_scheme {
                    let mut color_value: Option<String> = None;
                    if e.name().as_ref() == b"a:srgbClr" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                color_value =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    if let Some(name) = current_color_name.take() {
                        if color_value.is_some() {
                            theme.colors.push(ThemeColor {
                                name,
                                value: color_value,
                            });
                        }
                    }
                }

                if e.name().as_ref() == b"a:latin" {
                    let mut typeface = None;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"typeface" {
                            typeface = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                    if let Some(tf) = typeface {
                        if in_major_font {
                            font_scheme.major = Some(tf);
                        } else if in_minor_font {
                            font_scheme.minor = Some(tf);
                        }
                    }
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"a:clrScheme" => {
                    in_clr_scheme = false;
                }
                b"a:majorFont" => in_major_font = false,
                b"a:minorFont" => in_minor_font = false,
                _ => {
                    current_color_name = None;
                }
            },
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

    theme.fonts = font_scheme;
    Ok(theme)
}

pub fn parse_people_part(xml: &str, path: &str) -> Result<PeoplePart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut people = PeoplePart::new();
    people.span = Some(SourceSpan::new(path));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref().ends_with(b"person") {
                    let mut entry = PersonEntry {
                        person_id: None,
                        user_id: None,
                        display_name: None,
                        initials: None,
                    };
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let key = match key.iter().rposition(|b| *b == b':') {
                            Some(pos) => &key[pos + 1..],
                            None => key,
                        };
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key {
                            b"id" => entry.person_id = Some(val),
                            b"userId" | b"userID" => entry.user_id = Some(val),
                            b"displayName" | b"displayname" => entry.display_name = Some(val),
                            b"initials" => entry.initials = Some(val),
                            _ => {}
                        }
                    }
                    if entry.person_id.is_some()
                        || entry.user_id.is_some()
                        || entry.display_name.is_some()
                        || entry.initials.is_some()
                    {
                        people.people.push(entry);
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

    Ok(people)
}

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

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(pos) => &name[pos + 1..],
        None => name,
    }
}

pub fn parse_web_extension(xml: &str, path: &str) -> Result<WebExtension, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut ext = WebExtension::new();
    ext.span = Some(SourceSpan::new(path));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                match local {
                    b"webextension" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"id" || key == b"rId" || key == b"rid" {
                                ext.extension_id = Some(val);
                            }
                        }
                    }
                    b"storeReference" | b"storereference" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key {
                                b"store" => ext.store = Some(val),
                                b"storeType" | b"storetype" => ext.store_type = Some(val),
                                b"id" => ext.store_id = Some(val),
                                b"version" => ext.version = Some(val),
                                _ => {}
                            }
                        }
                    }
                    b"reference" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key {
                                b"id" => ext.reference_id = Some(val),
                                b"version" => ext.reference_version = Some(val),
                                b"store" => ext.store = Some(val),
                                b"storeType" | b"storetype" => ext.store_type = Some(val),
                                _ => {}
                            }
                        }
                    }
                    b"property" => {
                        let mut name = None;
                        let mut value = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key {
                                b"name" => name = Some(val),
                                b"value" | b"val" => value = Some(val),
                                _ => {}
                            }
                        }
                        if let (Some(name), Some(value)) = (name, value) {
                            ext.properties.push(WebExtensionProperty { name, value });
                        }
                    }
                    _ => {}
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

    Ok(ext)
}

pub fn parse_web_extension_taskpanes(
    xml: &str,
    path: &str,
) -> Result<Vec<WebExtensionTaskpane>, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut panes: Vec<WebExtensionTaskpane> = Vec::new();
    let mut current: Option<WebExtensionTaskpane> = None;

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"taskpane" {
                    let mut pane = WebExtensionTaskpane::new();
                    pane.span = Some(SourceSpan::new(path));
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key {
                            b"dockState" | b"dockstate" => pane.dock_state = Some(val),
                            b"visibility" => {
                                let v = val.eq_ignore_ascii_case("true") || val == "1";
                                pane.visibility = Some(v);
                            }
                            b"width" => pane.width = val.parse::<u32>().ok(),
                            b"height" => pane.height = val.parse::<u32>().ok(),
                            b"row" => pane.row = val.parse::<u32>().ok(),
                            b"column" => pane.column = val.parse::<u32>().ok(),
                            _ => {}
                        }
                    }
                    current = Some(pane);
                } else if local == b"webextensionref" {
                    if let Some(pane) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"id" || key == b"rid" || key == b"rId" {
                                pane.web_extension_ref = Some(val);
                            }
                        }
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"taskpane" {
                    let mut pane = WebExtensionTaskpane::new();
                    pane.span = Some(SourceSpan::new(path));
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key {
                            b"dockState" | b"dockstate" => pane.dock_state = Some(val),
                            b"visibility" => {
                                let v = val.eq_ignore_ascii_case("true") || val == "1";
                                pane.visibility = Some(v);
                            }
                            b"width" => pane.width = val.parse::<u32>().ok(),
                            b"height" => pane.height = val.parse::<u32>().ok(),
                            b"row" => pane.row = val.parse::<u32>().ok(),
                            b"column" => pane.column = val.parse::<u32>().ok(),
                            _ => {}
                        }
                    }
                    panes.push(pane);
                } else if local == b"webextensionref" {
                    if let Some(pane) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"id" || key == b"rid" || key == b"rId" {
                                pane.web_extension_ref = Some(val);
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"taskpane" {
                    if let Some(pane) = current.take() {
                        panes.push(pane);
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

    Ok(panes)
}

pub fn parse_vml_drawing(
    xml: &str,
    path: &str,
    rels: &Relationships,
) -> Result<(VmlDrawing, Vec<VmlShape>), ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    reader.config_mut().check_end_names = false;

    let mut drawing = VmlDrawing::new(path);
    drawing.span = Some(SourceSpan::new(path));
    let mut shapes: Vec<VmlShape> = Vec::new();
    let mut current: Option<VmlShape> = None;

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"shape" {
                    let mut shape = VmlShape::new();
                    shape.span = Some(SourceSpan::new(path));
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key {
                            b"id" | b"name" => shape.name = Some(val),
                            b"type" => shape.shape_type = Some(val),
                            b"style" => shape.style = Some(val),
                            b"filled" => {
                                shape.filled = Some(val == "t" || val == "true" || val == "1")
                            }
                            b"stroked" => {
                                shape.stroked = Some(val == "t" || val == "true" || val == "1")
                            }
                            _ => {}
                        }
                    }
                    current = Some(shape);
                } else if local == b"imagedata" {
                    if let Some(shape) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"id" || key == b"rid" || key == b"rId" {
                                shape.rel_id = Some(val.clone());
                                if let Some(rel) = rels.get(&val) {
                                    shape.image_target = Some(rel.target.clone());
                                }
                            }
                        }
                    }
                } else if local == b"textbox" {
                    if let Some(shape) = current.as_mut() {
                        let text = read_textbox_text(&mut reader)?;
                        if !text.is_empty() {
                            shape.text = Some(text);
                        }
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"shape" {
                    let mut shape = VmlShape::new();
                    shape.span = Some(SourceSpan::new(path));
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key {
                            b"id" | b"name" => shape.name = Some(val),
                            b"type" => shape.shape_type = Some(val),
                            b"style" => shape.style = Some(val),
                            b"filled" => {
                                shape.filled = Some(val == "t" || val == "true" || val == "1")
                            }
                            b"stroked" => {
                                shape.stroked = Some(val == "t" || val == "true" || val == "1")
                            }
                            _ => {}
                        }
                    }
                    shapes.push(shape);
                } else if local == b"imagedata" {
                    if let Some(shape) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"id" || key == b"rid" || key == b"rId" {
                                shape.rel_id = Some(val.clone());
                                if let Some(rel) = rels.get(&val) {
                                    shape.image_target = Some(rel.target.clone());
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"shape" {
                    if let Some(shape) = current.take() {
                        shapes.push(shape);
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

    Ok((drawing, shapes))
}

pub fn parse_drawingml_part(
    xml: &str,
    path: &str,
    rels: &Relationships,
) -> Result<(DrawingPart, Vec<Shape>), ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut part = DrawingPart::new(path);
    part.span = Some(SourceSpan::new(path));
    let mut shapes: Vec<Shape> = Vec::new();

    let mut buf = Vec::new();
    let mut current_name: Option<String> = None;
    let mut current_diagram_rel_ids: Vec<String> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                match local {
                    b"docPr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"name" {
                                current_name =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    b"blip" => {
                        let mut rel_id = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"embed" || key == b"link" || key == b"id" {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        if let Some(rel_id) = rel_id {
                            if let Some(rel) = rels.get(&rel_id) {
                                let mut shape = Shape::new(ShapeType::Picture);
                                shape.name = current_name.clone();
                                shape.relationship_id = Some(rel_id);
                                shape.media_target =
                                    Some(resolve_drawingml_target(path, &rel.target));
                                shape.span = Some(SourceSpan::new(path));
                                shapes.push(shape);
                            }
                        }
                    }
                    b"relIds" => {
                        current_diagram_rel_ids.clear();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"dm" || key == b"lo" || key == b"qs" || key == b"cs" {
                                current_diagram_rel_ids
                                    .push(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        if !current_diagram_rel_ids.is_empty() {
                            let mut related_targets = Vec::new();
                            for rel_id in &current_diagram_rel_ids {
                                if let Some(rel) = rels.get(rel_id) {
                                    related_targets
                                        .push(resolve_drawingml_target(path, &rel.target));
                                }
                            }
                            let mut shape = Shape::new(ShapeType::Custom);
                            shape.name = current_name.clone();
                            shape.relationship_id = current_diagram_rel_ids.first().cloned();
                            shape.related_targets = related_targets;
                            shape.span = Some(SourceSpan::new(path));
                            shapes.push(shape);
                        }
                        current_diagram_rel_ids.clear();
                    }
                    b"chart" => {
                        let mut rel_id = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"id" || key == b"rid" {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        if let Some(rel_id) = rel_id {
                            if let Some(rel) = rels.get(&rel_id) {
                                let mut shape = Shape::new(ShapeType::Chart);
                                shape.name = current_name.clone();
                                shape.relationship_id = Some(rel_id);
                                shape.media_target =
                                    Some(resolve_drawingml_target(path, &rel.target));
                                shape.span = Some(SourceSpan::new(path));
                                shapes.push(shape);
                            }
                        }
                    }
                    _ => {}
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

    Ok((part, shapes))
}

fn resolve_drawingml_target(path: &str, target: &str) -> String {
    if path.starts_with("word/") {
        normalize_docx_target(target)
    } else {
        Relationships::resolve_target(path, target)
    }
}

fn normalize_docx_target(target: &str) -> String {
    let mut t = target;
    while t.starts_with("../") {
        t = &t[3..];
    }
    if t.starts_with("./") {
        t = &t[2..];
    }
    if t.starts_with("word/") {
        t.to_string()
    } else {
        format!("word/{}", t.trim_start_matches('/'))
    }
}

fn read_textbox_text(reader: &mut Reader<&[u8]>) -> Result<String, ParseError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Text(t)) => {
                text.push_str(&t.unescape().unwrap_or_default());
            }
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"t" {
                    if let Ok(Event::Text(t)) = reader.read_event_into(&mut buf) {
                        text.push_str(&t.unescape().unwrap_or_default());
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                if local_name(&name_buf) == b"textbox" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "vml_textbox".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

pub fn parse_custom_xml_part(
    xml: &str,
    path: &str,
    size_bytes: u64,
) -> Result<CustomXmlPart, ParseError> {
    let mut part = CustomXmlPart::new(path, size_bytes);
    part.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut namespaces: HashSet<String> = HashSet::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                part.root_element = Some(String::from_utf8_lossy(e.name().as_ref()).to_string());
                for attr in e.attributes().flatten() {
                    let key = String::from_utf8_lossy(attr.key.as_ref());
                    if key.starts_with("xmlns") {
                        namespaces.insert(String::from_utf8_lossy(&attr.value).to_string());
                    }
                }
                break;
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

    part.namespaces = namespaces.into_iter().collect();
    Ok(part)
}

pub fn parse_signature(xml: &str, path: &str) -> Result<DigitalSignature, ParseError> {
    let mut sig = DigitalSignature::new();
    sig.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"ds:Signature" | b"Signature" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"Id" {
                            sig.signature_id =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"ds:SignatureMethod" | b"SignatureMethod" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"Algorithm" {
                            sig.signature_method =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"ds:DigestMethod" | b"DigestMethod" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"Algorithm" {
                            sig.digest_methods
                                .push(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"ds:X509SubjectName" | b"X509SubjectName" => {
                    if let Ok(text) = reader.read_text(e.name()) {
                        sig.signer = Some(text.to_string());
                    }
                }
                _ => {}
            },
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

    Ok(sig)
}

pub fn classify_media_type(path: &str) -> MediaType {
    let ext = path.rsplit('.').next().unwrap_or("").to_ascii_lowercase();
    match ext.as_str() {
        "jpg" | "jpeg" | "png" | "gif" | "bmp" | "tif" | "tiff" | "webp" => MediaType::Image,
        "mp3" | "wav" | "m4a" | "aac" | "ogg" => MediaType::Audio,
        "mp4" | "mov" | "avi" | "mkv" | "webm" => MediaType::Video,
        _ => MediaType::Other,
    }
}

pub fn legacy_extension_part(path: &str, size_bytes: u64) -> ExtensionPart {
    ExtensionPart::new(path, size_bytes, ExtensionPartKind::Legacy)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::relationships::Relationships;

    #[test]
    fn test_parse_theme_basic() {
        let xml = r#"
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
          <a:themeElements>
            <a:clrScheme name="Office">
              <a:dk1><a:srgbClr val="000000"/></a:dk1>
              <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
            </a:clrScheme>
            <a:fontScheme name="Office">
              <a:majorFont><a:latin typeface="Calibri"/></a:majorFont>
              <a:minorFont><a:latin typeface="Calibri Light"/></a:minorFont>
            </a:fontScheme>
          </a:themeElements>
        </a:theme>
        "#;
        let theme = parse_theme(xml, "theme/theme1.xml").expect("theme parse");
        assert_eq!(theme.name.as_deref(), Some("Office"));
        assert!(theme.colors.len() >= 2);
        assert_eq!(theme.fonts.major.as_deref(), Some("Calibri"));
    }

    #[test]
    fn test_parse_people_part() {
        let xml = r#"
        <ppl:people xmlns:ppl="http://schemas.openxmlformats.org/officeDocument/2006/relationships/people">
          <ppl:person ppl:id="p1" ppl:userId="user1" ppl:displayName="Alice" ppl:initials="A" />
          <ppl:person ppl:id="p2" ppl:userId="user2" ppl:displayName="Bob" />
        </ppl:people>
        "#;

        let people = parse_people_part(xml, "word/people.xml").expect("people");
        assert_eq!(people.people.len(), 2);
        assert_eq!(people.people[0].display_name.as_deref(), Some("Alice"));
        assert_eq!(people.people[1].display_name.as_deref(), Some("Bob"));
    }

    #[test]
    fn test_parse_web_extension() {
        let xml = r#"
        <we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11">
          <we:storeReference store="Store" storeType="OMEX" id="ext-1" version="1.0.0"/>
          <we:reference id="ref-1" version="1.0.0"/>
          <we:properties>
            <we:property name="foo" value="bar"/>
          </we:properties>
        </we:webextension>
        "#;

        let ext = parse_web_extension(xml, "word/webExtensions/webExtension1.xml").expect("ext");
        assert_eq!(ext.store.as_deref(), Some("Store"));
        assert_eq!(ext.store_type.as_deref(), Some("OMEX"));
        assert_eq!(ext.store_id.as_deref(), Some("ext-1"));
        assert_eq!(ext.reference_id.as_deref(), Some("ref-1"));
        assert_eq!(ext.properties.len(), 1);
    }

    #[test]
    fn test_parse_web_extension_taskpanes() {
        let xml = r#"
        <wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <wetp:taskpane dockState="right" visibility="true" width="300">
            <wetp:webextensionref r:id="rId1"/>
          </wetp:taskpane>
        </wetp:taskpanes>
        "#;

        let panes = parse_web_extension_taskpanes(xml, "word/webExtensions/taskpanes.xml")
            .expect("taskpanes");
        assert_eq!(panes.len(), 1);
        assert_eq!(panes[0].dock_state.as_deref(), Some("right"));
        assert_eq!(panes[0].web_extension_ref.as_deref(), Some("rId1"));
        assert_eq!(panes[0].width, Some(300));
    }

    #[test]
    fn test_parse_vml_drawing() {
        let xml = r##"
        <xml xmlns:v="urn:schemas-microsoft-com:vml"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <v:shape id="shape1" type="#_x0000_t75" style="width:100pt;height:50pt;">
            <v:imagedata r:id="rId1"/>
          </v:shape>
        </xml>
        "##;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels parse");
        let (drawing, shapes) = parse_vml_drawing(xml, "word/vmlDrawing1.vml", &rels).expect("vml");
        assert_eq!(drawing.path, "word/vmlDrawing1.vml");
        assert_eq!(shapes.len(), 1);
        assert_eq!(shapes[0].name.as_deref(), Some("shape1"));
        assert_eq!(shapes[0].image_target.as_deref(), Some("media/image1.png"));
    }

    #[test]
    fn test_parse_drawingml_part() {
        let xml = r#"
        <w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                   xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                   xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                   xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <wp:inline>
            <wp:docPr id="1" name="Image 1"/>
            <a:graphic>
              <a:graphicData>
                <a:blip r:embed="rId1"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
          <wp:inline>
            <wp:docPr id="2" name="Chart 1"/>
            <a:graphic>
              <a:graphicData>
                <c:chart r:id="rId2"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
          <wp:inline>
            <wp:docPr id="3" name="SmartArt 1"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <dgm:relIds r:dm="rId3" r:lo="rId4"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>
          <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="diagrams/layout1.xml"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels parse");
        let (_part, shapes) =
            parse_drawingml_part(xml, "word/drawings/drawing1.xml", &rels).expect("drawingml");
        assert_eq!(shapes.len(), 3);
        assert_eq!(shapes[0].name.as_deref(), Some("Image 1"));
        assert_eq!(
            shapes[0].media_target.as_deref(),
            Some("word/media/image1.png")
        );
        assert_eq!(shapes[1].shape_type, ShapeType::Chart);
        assert_eq!(
            shapes[1].media_target.as_deref(),
            Some("word/charts/chart1.xml")
        );
        assert_eq!(shapes[2].name.as_deref(), Some("SmartArt 1"));
        assert_eq!(shapes[2].shape_type, ShapeType::Custom);
        assert!(shapes[2]
            .related_targets
            .contains(&"word/diagrams/data1.xml".to_string()));
        assert!(shapes[2]
            .related_targets
            .contains(&"word/diagrams/layout1.xml".to_string()));
    }

    #[test]
    fn test_parse_custom_xml_part() {
        let xml = r#"<root xmlns="urn:example" xmlns:x="urn:x"><x:child/></root>"#;
        let part = parse_custom_xml_part(xml, "customXml/item1.xml", 42).expect("custom xml");
        assert_eq!(part.root_element.as_deref(), Some("root"));
        assert!(part.namespaces.iter().any(|n| n == "urn:example"));
    }

    #[test]
    fn test_build_relationship_graph() {
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="t1" Target="doc.xml"/>
          <Relationship Id="rId2" Type="t2" Target="http://example.com" TargetMode="External"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels parse");
        let graph =
            build_relationship_graph("word/document.xml", "word/_rels/document.xml.rels", &rels);
        assert_eq!(graph.relationships.len(), 2);
        assert!(graph
            .relationships
            .iter()
            .any(|r| matches!(r.target_mode, RelationshipTargetMode::External)));
    }
}
