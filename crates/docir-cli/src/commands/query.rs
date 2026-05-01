//! Query the IR with simple predicates.

use crate::commands::util::{parse_doc_format, parse_node_type, run_json_document_command};
use anyhow::Result;
use docir_app::ParserConfig;
use docir_core::ir::IrNode as IrNodeTrait;
use docir_core::query::Query;
use docir_core::types::NodeType;
use serde::Serialize;
use std::path::PathBuf;

#[derive(Debug, Clone)]
pub(crate) struct QueryFilters {
    pub(crate) node_type: Option<String>,
    pub(crate) contains: Option<String>,
    pub(crate) format: Option<String>,
    pub(crate) has_external_refs: Option<bool>,
    pub(crate) has_macros: Option<bool>,
}

#[derive(Debug, Serialize)]
struct QueryMatch {
    node_id: String,
    node_type: NodeType,
    location: Option<String>,
}

#[derive(Debug, Serialize)]
struct QueryResult {
    matches: Vec<QueryMatch>,
}

pub(crate) fn run_with_filters(
    input: PathBuf,
    filters: QueryFilters,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    run_json_document_command(input, parser_config, pretty, output, move |parsed| {
        let mut query = Query::new();

        if let Some(t) = filters.node_type {
            query.node_types.push(parse_node_type(&t)?);
        }
        if let Some(text) = filters.contains {
            query.text_contains = Some(text);
        }
        if let Some(fmt) = filters.format {
            query.format = Some(parse_doc_format(&fmt)?);
        }
        query.has_external_refs = filters.has_external_refs;
        query.has_macros = filters.has_macros;

        let matches = query
            .execute(parsed.store(), parsed.root_id())
            .into_iter()
            .filter_map(|id| {
                parsed.store().get(id).map(|node| QueryMatch {
                    node_id: id.to_string(),
                    node_type: node.node_type(),
                    location: node.source_span().map(|s| s.file_path.clone()),
                })
            })
            .collect();

        Ok(QueryResult { matches })
    })
}

#[cfg(test)]
mod tests {
    use super::{run_with_filters, QueryFilters};
    use docir_app::ParserConfig;
    use std::fs;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn fixture(name: &str) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../fixtures/ooxml")
            .join(name)
    }

    fn temp_file(name: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_query_{name}_{nanos}.json"))
    }

    #[test]
    fn query_run_with_filters_writes_matches_json() {
        let output = temp_file("filters");
        run_with_filters(
            fixture("minimal.docx"),
            QueryFilters {
                node_type: Some("Paragraph".to_string()),
                contains: Some("Hello".to_string()),
                format: Some("docx".to_string()),
                has_external_refs: None,
                has_macros: None,
            },
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("query run");
        let text = fs::read_to_string(&output).expect("query output");
        assert!(text.contains("matches"));
        let _ = fs::remove_file(output);
    }
}
