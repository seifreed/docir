//! HWP/HWPX part registry for coverage tracking.

use docir_core::types::DocumentFormat;

#[derive(Debug, Clone)]
pub struct PartSpec {
    pub format: DocumentFormat,
    pub pattern: &'static str,
    pub expected_parser: &'static str,
}

pub fn registry_for(format: DocumentFormat) -> Vec<PartSpec> {
    match format {
        DocumentFormat::Hwp => hwp_registry(),
        DocumentFormat::Hwpx => hwpx_registry(),
        _ => Vec::new(),
    }
}

pub fn matches_pattern(path: &str, pattern: &str) -> bool {
    if !pattern.contains('*') {
        return path == pattern;
    }

    let mut parts = pattern.splitn(2, '*');
    let prefix = parts.next().unwrap_or("");
    let suffix = parts.next().unwrap_or("");

    if !prefix.is_empty() && !path.starts_with(prefix) {
        return false;
    }
    if !suffix.is_empty() && !path.ends_with(suffix) {
        return false;
    }
    true
}

fn hwp_registry() -> Vec<PartSpec> {
    vec![
        PartSpec {
            format: DocumentFormat::Hwp,
            pattern: "FileHeader",
            expected_parser: "HwpParser::parse_header",
        },
        PartSpec {
            format: DocumentFormat::Hwp,
            pattern: "DocInfo",
            expected_parser: "HwpParser::parse_docinfo",
        },
        PartSpec {
            format: DocumentFormat::Hwp,
            pattern: "BodyText/Section*",
            expected_parser: "HwpParser::parse_bodytext",
        },
        PartSpec {
            format: DocumentFormat::Hwp,
            pattern: "BinData/*",
            expected_parser: "HwpParser::parse_bindata",
        },
        PartSpec {
            format: DocumentFormat::Hwp,
            pattern: "Scripts/*",
            expected_parser: "HwpParser::parse_scripts",
        },
        PartSpec {
            format: DocumentFormat::Hwp,
            pattern: "PrvText",
            expected_parser: "HwpParser::parse_preview_text",
        },
        PartSpec {
            format: DocumentFormat::Hwp,
            pattern: "PrvImage",
            expected_parser: "HwpParser::parse_preview_image",
        },
        PartSpec {
            format: DocumentFormat::Hwp,
            pattern: "HwpSummaryInformation",
            expected_parser: "HwpParser::parse_summary",
        },
        PartSpec {
            format: DocumentFormat::Hwp,
            pattern: "SummaryInformation",
            expected_parser: "HwpParser::parse_summary",
        },
    ]
}

fn hwpx_registry() -> Vec<PartSpec> {
    vec![
        PartSpec {
            format: DocumentFormat::Hwpx,
            pattern: "mimetype",
            expected_parser: "HwpxParser::parse_mimetype",
        },
        PartSpec {
            format: DocumentFormat::Hwpx,
            pattern: "META-INF/container.xml",
            expected_parser: "HwpxParser::parse_container",
        },
        PartSpec {
            format: DocumentFormat::Hwpx,
            pattern: "version.xml",
            expected_parser: "HwpxParser::parse_version",
        },
        PartSpec {
            format: DocumentFormat::Hwpx,
            pattern: "Contents/content.hpf",
            expected_parser: "HwpxParser::parse_package_info",
        },
        PartSpec {
            format: DocumentFormat::Hwpx,
            pattern: "Contents/section*.xml",
            expected_parser: "HwpxParser::parse_section",
        },
        PartSpec {
            format: DocumentFormat::Hwpx,
            pattern: "Contents/header*.xml",
            expected_parser: "HwpxParser::parse_header",
        },
        PartSpec {
            format: DocumentFormat::Hwpx,
            pattern: "Contents/footer*.xml",
            expected_parser: "HwpxParser::parse_footer",
        },
        PartSpec {
            format: DocumentFormat::Hwpx,
            pattern: "Contents/masterPage*.xml",
            expected_parser: "HwpxParser::parse_master_page",
        },
        PartSpec {
            format: DocumentFormat::Hwpx,
            pattern: "BinData/*",
            expected_parser: "HwpxParser::parse_bindata",
        },
    ]
}
