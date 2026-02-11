//! Run rule engine on a document.

use crate::commands::util::{build_parser, write_json_output};
use anyhow::Result;
use docir_parser::ParserConfig;
use docir_rules::{RuleEngine, RuleProfile};
use std::fs::File;
use std::path::PathBuf;

pub fn run(
    input: PathBuf,
    pretty: bool,
    output: Option<PathBuf>,
    profile_path: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let parser = build_parser(parser_config);
    let parsed = parser.parse_file(&input)?;

    let profile = if let Some(path) = profile_path {
        let file = File::open(path)?;
        serde_json::from_reader(file)?
    } else {
        RuleProfile::default()
    };
    let engine = RuleEngine::with_default_rules();
    let report = engine.run_with_profile(&parsed.store, parsed.root_id, &profile);

    write_json_output(&report, pretty, output)
}
