//! Grep-like semantic search (text contains).

use anyhow::{bail, Result};
use docir_app::ParserConfig;
use std::path::PathBuf;

use crate::cli::PrettyOutputOpts;
use crate::commands::query::{run_with_filters, QueryFilters};

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    pattern: String,
    node_type: Option<String>,
    format: Option<String>,
    opts: PrettyOutputOpts,
    parser_config: &ParserConfig,
) -> Result<()> {
    if pattern.trim().is_empty() {
        bail!("Pattern must not be empty");
    }
    run_with_filters(
        input,
        QueryFilters {
            node_type,
            contains: Some(pattern),
            format,
            has_external_refs: None,
            has_macros: None,
        },
        opts,
        parser_config,
    )
}

#[cfg(test)]
mod tests {
    use super::run;
    use crate::cli::PrettyOutputOpts;
    use crate::test_support;
    use docir_app::ParserConfig;
    use std::fs;

    #[test]
    fn grep_run_rejects_empty_pattern() {
        let err = run(
            test_support::fixture("minimal.docx"),
            "   ".to_string(),
            None,
            None,
            PrettyOutputOpts {
                pretty: false,
                output: None,
            },
            &ParserConfig::default(),
        )
        .expect_err("empty grep pattern should fail");
        assert!(err.to_string().contains("Pattern must not be empty"));
    }

    #[test]
    fn grep_run_outputs_results_file() {
        let output = test_support::temp_file("results", "json");
        run(
            test_support::fixture("minimal.docx"),
            "Hello".to_string(),
            None,
            None,
            PrettyOutputOpts {
                pretty: true,
                output: Some(output.clone()),
            },
            &ParserConfig::default(),
        )
        .expect("grep should run");
        let text = fs::read_to_string(&output).expect("grep output");
        assert!(text.contains("matches"));
        let _ = fs::remove_file(output);
    }
}
