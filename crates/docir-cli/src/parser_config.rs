//! Parser configuration wiring from CLI flags.

use crate::cli::Cli;
use docir_app::ParserConfig;

pub(crate) fn build_parser_config(cli: &Cli) -> ParserConfig {
    let mut config = ParserConfig::default();
    apply_zip_overrides(cli, &mut config);
    apply_odf_overrides(cli, &mut config);
    apply_hwp_overrides(cli, &mut config);
    copy_if_some(cli.max_input_size, &mut config.max_input_size);
    set_if(cli.metrics, &mut config.enable_metrics);
    clear_if(cli.no_hashes, &mut config.compute_hashes);
    config
}

fn apply_zip_overrides(cli: &Cli, config: &mut ParserConfig) {
    copy_if_some(
        cli.zip_max_total_size,
        &mut config.zip_config.max_total_size,
    );
    copy_if_some(cli.zip_max_file_size, &mut config.zip_config.max_file_size);
    copy_if_some(
        cli.zip_max_file_count,
        &mut config.zip_config.max_file_count,
    );
    copy_if_some(
        cli.zip_max_compression_ratio,
        &mut config.zip_config.max_compression_ratio,
    );
    copy_if_some(
        cli.zip_max_path_depth,
        &mut config.zip_config.max_path_depth,
    );
}

fn apply_odf_overrides(cli: &Cli, config: &mut ParserConfig) {
    set_if(cli.odf_fast, &mut config.odf.force_fast);
    copy_if_some(
        cli.odf_fast_threshold_bytes,
        &mut config.odf.fast_threshold_bytes,
    );
    copy_if_some(cli.odf_fast_sample_rows, &mut config.odf.fast_sample_rows);
    copy_if_some(cli.odf_fast_sample_cols, &mut config.odf.fast_sample_cols);
    copy_if_some(cli.odf_max_cells.map(non_zero), &mut config.odf.max_cells);
    copy_if_some(cli.odf_max_rows.map(non_zero), &mut config.odf.max_rows);
    copy_if_some(
        cli.odf_max_paragraphs.map(non_zero),
        &mut config.odf.max_paragraphs,
    );
    copy_if_some(cli.odf_max_bytes.map(non_zero), &mut config.odf.max_bytes);
    set_if(cli.odf_parallel_sheets, &mut config.odf.parallel_sheets);
    copy_if_some(
        cli.odf_parallel_max_threads.map(Some),
        &mut config.odf.parallel_max_threads,
    );
    copy_if_some(cli.odf_password.clone().map(Some), &mut config.odf.password);
}

fn apply_hwp_overrides(cli: &Cli, config: &mut ParserConfig) {
    set_if(
        cli.hwp_force_parse_encrypted,
        &mut config.hwp.force_parse_encrypted,
    );
    copy_if_some(cli.hwp_password.clone().map(Some), &mut config.hwp.password);
    set_if(cli.hwp_dump_streams, &mut config.hwp.dump_streams);
}

fn copy_if_some<T>(value: Option<T>, target: &mut T) {
    if let Some(value) = value {
        *target = value;
    }
}

fn set_if(flag: bool, target: &mut bool) {
    if flag {
        *target = true;
    }
}

fn clear_if(flag: bool, target: &mut bool) {
    if flag {
        *target = false;
    }
}

fn non_zero(value: u64) -> Option<u64> {
    (value != 0).then_some(value)
}

#[cfg(test)]
mod tests {
    use super::build_parser_config;
    use crate::Cli;
    use clap::Parser;

    #[test]
    fn build_parser_config_disables_hashes_when_requested() {
        let cli = Cli::parse_from(["docir", "--no-hashes", "inventory", "sample.docx"]);
        let config = build_parser_config(&cli);
        assert!(!config.compute_hashes);
    }

    #[test]
    fn build_parser_config_keeps_hashes_enabled_by_default() {
        let cli = Cli::parse_from(["docir", "inventory", "sample.docx"]);
        let config = build_parser_config(&cli);
        assert!(config.compute_hashes);
    }
}
