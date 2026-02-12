//! Python bindings for docir.

use docir_app::DocirApp;
use docir_app::ParserConfig;
use docir_core::ir::IrNode as IrNodeTrait;
use docir_core::query::Query;
use docir_core::types::{
    parse_document_format as parse_core_document_format, parse_node_type as parse_core_node_type,
    DocumentFormat, NodeType,
};
use docir_core::visitor::{NodeCounter, PreOrderWalker};
use docir_rules::RuleProfile;
use docir_serialization::{IrSerializer, JsonSerializer};
use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use serde_json::json;
use std::path::Path;

#[pyfunction]
fn parse_json(path: String, pretty: Option<bool>) -> PyResult<String> {
    let app = DocirApp::new(ParserConfig::default());
    let parsed = app
        .parse_file(Path::new(&path))
        .map_err(|e| PyValueError::new_err(e.to_string()))?;
    let serializer = if pretty.unwrap_or(false) {
        JsonSerializer::pretty()
    } else {
        JsonSerializer::new()
    };
    serializer
        .serialize_to_string(parsed.store(), parsed.root_id())
        .map_err(|e| PyValueError::new_err(e.to_string()))
}

#[pyfunction]
fn rules(path: String, profile_json: Option<String>, pretty: Option<bool>) -> PyResult<String> {
    let app = DocirApp::new(ParserConfig::default());
    let parsed = app
        .parse_file(Path::new(&path))
        .map_err(|e| PyValueError::new_err(e.to_string()))?;

    let profile = if let Some(text) = profile_json {
        serde_json::from_str(&text).map_err(|e| PyValueError::new_err(e.to_string()))?
    } else {
        RuleProfile::default()
    };

    let report = app.run_rules(&parsed, &profile);

    if pretty.unwrap_or(false) {
        serde_json::to_string_pretty(&report).map_err(|e| PyValueError::new_err(e.to_string()))
    } else {
        serde_json::to_string(&report).map_err(|e| PyValueError::new_err(e.to_string()))
    }
}

#[pyfunction]
fn query(
    path: String,
    node_type: Option<String>,
    contains: Option<String>,
    format: Option<String>,
    has_external_refs: Option<bool>,
    has_macros: Option<bool>,
    pretty: Option<bool>,
) -> PyResult<String> {
    let app = DocirApp::new(ParserConfig::default());
    let parsed = app
        .parse_file(Path::new(&path))
        .map_err(|e| PyValueError::new_err(e.to_string()))?;

    let mut q = Query::new();
    if let Some(t) = node_type {
        q.node_types.push(parse_node_type(&t)?);
    }
    if let Some(text) = contains {
        q.text_contains = Some(text);
    }
    if let Some(fmt) = format {
        q.format = Some(parse_doc_format(&fmt)?);
    }
    q.has_external_refs = has_external_refs;
    q.has_macros = has_macros;

    let matches: Vec<_> = q
        .execute(parsed.store(), parsed.root_id())
        .into_iter()
        .filter_map(|id| {
            parsed.store().get(id).map(|node| {
                json!({
                    "node_id": id.to_string(),
                    "node_type": format!("{:?}", node.node_type()),
                    "location": node.source_span().map(|s| s.file_path.clone()),
                })
            })
        })
        .collect();

    let value = json!({"matches": matches});
    if pretty.unwrap_or(false) {
        serde_json::to_string_pretty(&value).map_err(|e| PyValueError::new_err(e.to_string()))
    } else {
        serde_json::to_string(&value).map_err(|e| PyValueError::new_err(e.to_string()))
    }
}

#[pyfunction]
fn summary(path: String) -> PyResult<String> {
    let app = DocirApp::new(ParserConfig::default());
    let parsed = app
        .parse_file(Path::new(&path))
        .map_err(|e| PyValueError::new_err(e.to_string()))?;

    let doc = parsed
        .document()
        .ok_or_else(|| PyValueError::new_err("Document root not found"))?;

    let mut counter = NodeCounter::new();
    let mut walker = PreOrderWalker::new(parsed.store(), parsed.root_id());
    let _ = walker.walk(&mut counter);

    let mut lines = Vec::new();
    lines.push(format!("Format: {}", doc.format.display_name()));
    lines.push(format!("Nodes: {}", parsed.store().len()));

    let mut counts: Vec<_> = counter.counts.iter().collect();
    counts.sort_by_key(|(_, v)| std::cmp::Reverse(*v));
    for (node_type, count) in counts {
        if *count > 0 {
            lines.push(format!("{}: {}", node_type, count));
        }
    }

    lines.push(format!("Threat Level: {}", doc.security.threat_level));
    lines.push(format!(
        "Macros: {}",
        if doc.security.macro_project.is_some() {
            "yes"
        } else {
            "no"
        }
    ));
    lines.push(format!("OLE Objects: {}", doc.security.ole_objects.len()));
    lines.push(format!(
        "External References: {}",
        doc.security.external_refs.len()
    ));
    lines.push(format!("DDE Fields: {}", doc.security.dde_fields.len()));

    Ok(lines.join("\n"))
}

fn parse_node_type(input: &str) -> PyResult<NodeType> {
    parse_core_node_type(input).map_err(|e| PyValueError::new_err(e.to_string()))
}

fn parse_doc_format(input: &str) -> PyResult<DocumentFormat> {
    parse_core_document_format(input).map_err(|e| PyValueError::new_err(e.to_string()))
}

#[pymodule]
fn docir(_py: Python<'_>, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(parse_json, m)?)?;
    m.add_function(wrap_pyfunction!(rules, m)?)?;
    m.add_function(wrap_pyfunction!(query, m)?)?;
    m.add_function(wrap_pyfunction!(summary, m)?)?;
    Ok(())
}
