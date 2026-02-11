//! Run rule engine on a document.

use crate::commands::util::{build_app, write_json_output};
use anyhow::Result;
use docir_parser::ParserConfig;
use docir_rules::RuleProfile;
use std::fs::File;
use std::path::PathBuf;

pub fn run(
    input: PathBuf,
    pretty: bool,
    output: Option<PathBuf>,
    profile_path: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let app = build_app(parser_config);
    let parsed = app.parse_file(&input)?;

    let profile = if let Some(path) = profile_path {
        let file = File::open(path)?;
        serde_json::from_reader(file)?
    } else {
        RuleProfile::default()
    };
    let report = app.run_rules(&parsed, &profile);

    write_json_output(&report, pretty, output)
}
