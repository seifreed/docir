//! Query the IR with simple predicates.

use crate::cli::PrettyOutputOpts;
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
    opts: PrettyOutputOpts,
    parser_config: &ParserConfig,
) -> Result<()> {
    let PrettyOutputOpts { pretty, output } = opts;
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
    use crate::cli::PrettyOutputOpts;
    use crate::test_support;
    use docir_app::ParserConfig;
    use std::fs;

    #[test]
    fn query_run_with_filters_writes_matches_json() {
        let output = test_support::temp_file("filters", "json");
        run_with_filters(
            test_support::fixture("minimal.docx"),
            QueryFilters {
                node_type: Some("Paragraph".to_string()),
                contains: Some("Hello".to_string()),
                format: Some("docx".to_string()),
                has_external_refs: None,
                has_macros: None,
            },
            PrettyOutputOpts {
                pretty: true,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("query run");
        let text = fs::read_to_string(&output).expect("query output");
        assert!(text.contains("matches"));
        let _ = fs::remove_file(output);
    }
}
