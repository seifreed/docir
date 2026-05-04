//! Run rule engine on a document.

use crate::cli::PrettyOutputOpts;
use crate::commands::util::run_json_app_command;
use anyhow::Result;
use docir_app::{ParserConfig, RuleProfile};
use std::fs::File;
use std::path::PathBuf;

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    opts: PrettyOutputOpts,
    profile_path: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let PrettyOutputOpts { pretty, output } = opts;
    run_json_app_command(parser_config, pretty, output, move |app| {
        let parsed = app.parse_file(&input)?;
        let profile = if let Some(path) = profile_path {
            let file = File::open(path)?;
            serde_json::from_reader(file)?
        } else {
            RuleProfile::default()
        };
        Ok(app.run_rules(&parsed, &profile))
    })
}
