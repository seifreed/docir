//! # docir CLI
//!
//! Command-line interface for the docir document analysis toolkit.

mod cli;
mod commands;
mod parser_config;

use anyhow::Result;
use clap::Parser;

pub(crate) use cli::{Cli, Commands, CoverageExportFormat, CoverageExportMode, OutputFormat};
pub(crate) use parser_config::build_parser_config;

fn main() -> Result<()> {
    env_logger::Builder::from_env(env_logger::Env::default().default_filter_or("warn")).init();

    let cli = Cli::parse();
    let parser_config = build_parser_config(&cli);
    commands::dispatch::run(cli, &parser_config)
}
