use anyhow::Result;
use docir_app::ParserConfig;
use std::path::PathBuf;

pub(crate) fn cmd_inspect_metadata(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::inspect_metadata::run(input, json, pretty, output, parser_config)
}

pub(crate) fn cmd_inspect_sheet_records(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::inspect_sheet_records::run(input, json, pretty, output, parser_config)
}

pub(crate) fn cmd_inspect_slide_records(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::inspect_slide_records::run(input, json, pretty, output, parser_config)
}

pub(crate) fn cmd_inspect_directory(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::inspect_directory::run(input, json, pretty, output, parser_config)
}

pub(crate) fn cmd_inspect_sectors(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    super::inspect_sectors::run(input, json, pretty, output, parser_config)
}
