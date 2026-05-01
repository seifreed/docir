//! Python bindings for docir.

use docir_app::RuleProfile;
use docir_app::{
    DocirApp, ExportDocumentRef, ParsedDocument, ParserConfig, Phase0ArtifactManifestExport,
    Phase0VbaExport,
};
use docir_core::ir::IrNode as IrNodeTrait;
use docir_core::query::Query;
use docir_core::types::{
    parse_document_format as parse_core_document_format, parse_node_type as parse_core_node_type,
    DocumentFormat, NodeType,
};
use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use serde_json::json;
use std::path::Path;

#[pyfunction]
fn parse_json(path: String, pretty: Option<bool>) -> PyResult<String> {
    let (app, parsed) = build_app_and_parse(&path)?;
    app.serialize_json(&parsed, pretty.unwrap_or(false))
        .map_err(to_py_value_error)
}

#[pyfunction]
fn rules(path: String, profile_json: Option<String>, pretty: Option<bool>) -> PyResult<String> {
    let (app, parsed) = build_app_and_parse(&path)?;

    let profile = if let Some(text) = profile_json {
        serde_json::from_str(&text).map_err(to_py_value_error)?
    } else {
        RuleProfile::default()
    };

    let report = app.run_rules(&parsed, &profile);
    to_json_string(&report, pretty.unwrap_or(false))
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
    let (_, parsed) = build_app_and_parse(&path)?;

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
    to_json_string(&value, pretty.unwrap_or(false))
}

#[pyfunction]
fn summary(path: String) -> PyResult<String> {
    let (app, parsed) = build_app_and_parse(&path)?;
    app.format_summary(&parsed, Some(&path))
        .ok_or_else(|| PyValueError::new_err("Could not build document summary"))
}

#[pyfunction]
fn inventory(path: String, pretty: Option<bool>) -> PyResult<String> {
    let (app, parsed) = build_app_and_parse(&path)?;
    let inventory = app.build_inventory(&parsed);
    let export = Phase0ArtifactManifestExport::from_inventory(
        &inventory,
        ExportDocumentRef::new(path.clone(), parsed.format().extension(), Some(path)),
    );
    to_json_string(&export, pretty.unwrap_or(false))
}

#[pyfunction]
fn recognize_vba(
    path: String,
    include_source: Option<bool>,
    pretty: Option<bool>,
) -> PyResult<String> {
    let mut config = ParserConfig::default();
    if include_source.unwrap_or(false) {
        config.extract_macro_source = true;
    }
    let app = DocirApp::new(config);
    let parsed = app
        .parse_file(Path::new(&path))
        .map_err(to_py_value_error)?;
    let report = app.build_vba_recognition(&parsed, include_source.unwrap_or(false));
    let export = Phase0VbaExport::from_report(
        &report,
        ExportDocumentRef::new(path.clone(), parsed.format().extension(), Some(path)),
    );
    to_json_string(&export, pretty.unwrap_or(false))
}

fn parse_node_type(input: &str) -> PyResult<NodeType> {
    parse_core_node_type(input).map_err(to_py_value_error)
}

fn parse_doc_format(input: &str) -> PyResult<DocumentFormat> {
    parse_core_document_format(input).map_err(to_py_value_error)
}

fn build_app_and_parse(path: &str) -> PyResult<(DocirApp, ParsedDocument)> {
    let app = DocirApp::new(ParserConfig::default());
    let parsed = app.parse_file(Path::new(path)).map_err(to_py_value_error)?;
    Ok((app, parsed))
}

fn to_json_string<T: serde::Serialize>(value: &T, pretty: bool) -> PyResult<String> {
    if pretty {
        serde_json::to_string_pretty(value).map_err(to_py_value_error)
    } else {
        serde_json::to_string(value).map_err(to_py_value_error)
    }
}

fn to_py_value_error<E: std::fmt::Display>(err: E) -> PyErr {
    PyValueError::new_err(err.to_string())
}

#[pymodule]
fn docir(_py: Python<'_>, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(parse_json, m)?)?;
    m.add_function(wrap_pyfunction!(rules, m)?)?;
    m.add_function(wrap_pyfunction!(query, m)?)?;
    m.add_function(wrap_pyfunction!(summary, m)?)?;
    m.add_function(wrap_pyfunction!(inventory, m)?)?;
    m.add_function(wrap_pyfunction!(recognize_vba, m)?)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::Value;
    use std::sync::Once;

    static PYTHON_INIT: Once = Once::new();

    fn init_python() {
        PYTHON_INIT.call_once(pyo3::prepare_freethreaded_python);
    }

    fn fixture(path: &str) -> String {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../fixtures")
            .join(path)
            .to_string_lossy()
            .to_string()
    }

    #[test]
    fn parse_json_emits_ir_json() {
        let docx = fixture("ooxml/minimal.docx");
        let text = parse_json(docx, Some(false)).expect("parse_json should succeed");
        let parsed: Value = serde_json::from_str(&text).expect("valid JSON output");
        assert!(parsed.is_object());
        assert!(parsed
            .as_object()
            .map(|map| !map.is_empty())
            .unwrap_or(false));
    }

    #[test]
    fn parse_json_pretty_emits_multiline_json() {
        let docx = fixture("ooxml/minimal.docx");
        let text = parse_json(docx, Some(true)).expect("parse_json pretty should succeed");
        assert!(text.contains('\n'));
        let parsed: Value = serde_json::from_str(&text).expect("valid JSON output");
        assert!(parsed.is_object());
    }

    #[test]
    fn rules_emits_report_json() {
        let docx = fixture("ooxml/minimal.docx");
        let text = rules(docx, None, Some(false)).expect("rules should succeed");
        let parsed: Value = serde_json::from_str(&text).expect("valid JSON output");
        assert!(parsed.get("findings").is_some());
    }

    #[test]
    fn rules_accepts_profile_json_and_pretty_output() {
        let docx = fixture("ooxml/minimal.docx");
        let profile = r#"{"disabled_rules":["burst-formula-length"]}"#.to_string();
        let text = rules(docx, Some(profile), Some(true)).expect("rules should parse profile");
        assert!(text.contains('\n'));
        let parsed: Value = serde_json::from_str(&text).expect("valid JSON output");
        assert!(parsed.get("findings").is_some());
    }

    #[test]
    fn rules_rejects_invalid_profile_json() {
        init_python();
        let docx = fixture("ooxml/minimal.docx");
        let err = rules(docx, Some("{".to_string()), Some(false)).expect_err("must fail");
        assert!(!err.to_string().is_empty());
    }

    #[test]
    fn query_emits_matches_array() {
        let docx = fixture("ooxml/minimal.docx");
        let text = query(
            docx,
            Some("document".to_string()),
            None,
            Some("docx".to_string()),
            None,
            None,
            Some(false),
        )
        .expect("query should succeed");
        let parsed: Value = serde_json::from_str(&text).expect("valid JSON output");
        assert!(parsed.get("matches").is_some());
    }

    #[test]
    fn query_supports_contains_and_boolean_filters() {
        let docx = fixture("ooxml/minimal.docx");
        let text = query(
            docx,
            Some("paragraph".to_string()),
            Some("Hello".to_string()),
            None,
            Some(false),
            Some(false),
            Some(false),
        )
        .expect("query with filters should succeed");
        let parsed: Value = serde_json::from_str(&text).expect("valid JSON output");
        let matches = parsed
            .get("matches")
            .and_then(|v| v.as_array())
            .expect("matches array");
        let _ = matches;
    }

    #[test]
    fn query_rejects_invalid_format_input() {
        init_python();
        let docx = fixture("ooxml/minimal.docx");
        let err = query(
            docx,
            None,
            None,
            Some("not-a-real-format".to_string()),
            None,
            None,
            Some(false),
        )
        .expect_err("invalid format should fail");
        assert!(!err.to_string().is_empty());
    }

    #[test]
    fn summary_includes_expected_headings() {
        let docx = fixture("ooxml/minimal.docx");
        let text = summary(docx).expect("summary should succeed");
        assert!(text.contains("Format:"));
        assert!(text.contains("Nodes:"));
    }

    #[test]
    fn summary_includes_macros_line() {
        let hwp = fixture("hwp/minimal.hwp");
        let text = summary(hwp).expect("summary should succeed for hwp");
        assert!(text.contains("Macros:"));
    }

    #[test]
    fn inventory_emits_inventory_json() {
        let docx = fixture("ooxml/minimal.docx");
        let text = inventory(docx, Some(false)).expect("inventory should succeed");
        let parsed: Value = serde_json::from_str(&text).expect("valid JSON output");
        assert!(parsed.get("schema_version").is_some());
        assert!(parsed.get("artifacts").is_some());
    }

    #[test]
    fn recognize_vba_emits_report_json() {
        let docx = fixture("ooxml/minimal.docx");
        let text = recognize_vba(docx, Some(false), Some(false)).expect("recognize_vba");
        let parsed: Value = serde_json::from_str(&text).expect("valid JSON output");
        assert!(parsed.get("vba").is_some());
        assert!(parsed.get("schema_version").is_some());
    }

    #[test]
    fn parse_node_type_rejects_invalid_type() {
        assert!(parse_node_type("not-a-real-node-type").is_err());
    }

    #[test]
    fn parse_doc_format_rejects_invalid_format() {
        assert!(parse_doc_format("not-a-real-format").is_err());
    }

    #[test]
    fn to_json_string_supports_pretty() {
        let compact = to_json_string(&json!({"k":"v"}), false).expect("compact");
        let pretty = to_json_string(&json!({"k":"v"}), true).expect("pretty");
        assert!(!compact.contains('\n'));
        assert!(pretty.contains('\n'));
    }

    #[test]
    fn docir_module_registers_public_functions() {
        init_python();
        Python::with_gil(|py| {
            let module = PyModule::new(py, "docir").expect("module creation");
            docir(py, &module).expect("module registration");
            assert!(module.getattr("parse_json").is_ok());
            assert!(module.getattr("rules").is_ok());
            assert!(module.getattr("query").is_ok());
            assert!(module.getattr("summary").is_ok());
        });
    }
}
