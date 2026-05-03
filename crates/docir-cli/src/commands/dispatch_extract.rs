use anyhow::Result;
use docir_app::ParserConfig;
use std::path::PathBuf;

use super::extract_artifacts::ExtractArtifactsOptions;

pub(crate) fn cmd_extract_links(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::extract_links::run(input, json, pretty, output, parser_config)
}

pub(crate) fn cmd_extract_flash(
    input: PathBuf,
    out: Option<PathBuf>,
    overwrite: bool,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::extract_flash::run(input, out, overwrite, json, pretty, output, parser_config)
}

pub(crate) fn cmd_extract_vba(
    input: PathBuf,
    out: PathBuf,
    overwrite: bool,
    best_effort: bool,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::extract_vba::run(input, out, overwrite, best_effort, parser_config)
}

pub(crate) fn cmd_extract_artifacts(
    input: PathBuf,
    out: PathBuf,
    options: ExtractArtifactsOptions,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::extract_artifacts::run(input, out, options, parser_config)
}

pub(crate) fn cmd_extract(
    input: PathBuf,
    node_id: Vec<String>,
    node_type: Option<String>,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::extract::run(input, node_id, node_type, pretty, output, parser_config)
}
