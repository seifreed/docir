//! Grep-like semantic search (text contains).

use anyhow::{bail, Result};
use docir_app::ParserConfig;
use std::path::PathBuf;

use crate::commands::query::{run_with_filters, QueryFilters};

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    pattern: String,
    node_type: Option<String>,
    format: Option<String>,
    pretty: bool,
    output: Option<PathBuf>,
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
        pretty,
        output,
        parser_config,
    )
}

#[cfg(test)]
mod tests {
    use super::run;
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
        std::env::temp_dir().join(format!("docir_cli_grep_{name}_{nanos}.json"))
    }

    #[test]
    fn grep_run_rejects_empty_pattern() {
        let err = run(
            fixture("minimal.docx"),
            "   ".to_string(),
            None,
            None,
            false,
            None,
            &ParserConfig::default(),
        )
        .expect_err("empty grep pattern should fail");
        assert!(err.to_string().contains("Pattern must not be empty"));
    }

    #[test]
    fn grep_run_outputs_results_file() {
        let output = temp_file("results");
        run(
            fixture("minimal.docx"),
            "Hello".to_string(),
            None,
            None,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("grep should run");
        let text = fs::read_to_string(&output).expect("grep output");
        assert!(text.contains("matches"));
        let _ = fs::remove_file(output);
    }
}
