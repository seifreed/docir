//! Inspect classic OLE metadata property sets.

use anyhow::Result;
use docir_app::{inspect_metadata_path, MetadataInspection, ParserConfig};
use std::path::PathBuf;

use crate::commands::util::{push_bullet_line, push_labeled_line, run_dual_output};

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let metadata = inspect_metadata_path(&input, parser_config)?;
    run_dual_output(
        &metadata,
        "metadata",
        json,
        pretty,
        output,
        format_metadata_text,
    )
}

fn format_metadata_text(metadata: &MetadataInspection) -> String {
    let mut out = String::new();
    push_labeled_line(&mut out, 0, "Container", &metadata.container);
    push_labeled_line(&mut out, 0, "Sections", metadata.section_count);
    for section in &metadata.sections {
        out.push('\n');
        push_labeled_line(&mut out, 0, "Section", &section.name);
        push_labeled_line(&mut out, 2, "Path", &section.path);
        push_labeled_line(&mut out, 2, "Properties", section.property_count);
        for property in &section.properties {
            push_bullet_line(
                &mut out,
                4,
                &property.name,
                property.display_value.as_deref().unwrap_or(&property.value),
            );
            push_labeled_line(&mut out, 6, "Type", &property.value_type);
            if let Some(display_value) = property.display_value.as_deref() {
                push_labeled_line(&mut out, 6, "Raw", &property.value);
                push_labeled_line(&mut out, 6, "Display", display_value);
            }
            push_labeled_line(&mut out, 6, "Id", property.id);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::{format_metadata_text, run};
    use crate::test_support;
    use docir_app::test_support::{
        build_test_cfb, build_test_property_set_stream, TestPropertyValue,
    };
    use docir_app::{MetadataInspection, MetadataProperty, MetadataSection, ParserConfig};
    use std::fs;

    const SUMMARY_INFO_STREAM: &str = "\u{0005}SummaryInformation";
    const DOC_SUMMARY_INFO_STREAM: &str = "\u{0005}DocumentSummaryInformation";

    #[test]
    fn inspect_metadata_run_writes_json() {
        let input = test_support::temp_file("legacy", "doc");
        let output = test_support::temp_file("legacy", "json");
        let summary = build_test_property_set_stream(&[
            (2, TestPropertyValue::Str("Specimen")),
            (3, TestPropertyValue::Str("Specimen subject")),
            (4, TestPropertyValue::Str("Analyst")),
            (5, TestPropertyValue::Str("macro,ole")),
            (8, TestPropertyValue::Str("Responder")),
            (10, TestPropertyValue::I64(3600)),
            (11, TestPropertyValue::FileTime(10_000_000)),
            (14, TestPropertyValue::I32(7)),
            (15, TestPropertyValue::I32(321)),
            (18, TestPropertyValue::Str("Microsoft Excel")),
            (19, TestPropertyValue::U32(1)),
        ]);
        let doc_summary = build_test_property_set_stream(&[
            (2, TestPropertyValue::Str("Malware triage")),
            (3, TestPropertyValue::Str("Screen")),
            (4, TestPropertyValue::I32(2048)),
            (14, TestPropertyValue::WStr("Analyst")),
            (15, TestPropertyValue::WStr("ACME")),
            (26, TestPropertyValue::WStr("application/vnd.ms-excel")),
            (27, TestPropertyValue::WStr("final")),
            (28, TestPropertyValue::WStr("en-US")),
        ]);
        fs::write(
            &input,
            build_test_cfb(&[
                (SUMMARY_INFO_STREAM, &summary),
                (DOC_SUMMARY_INFO_STREAM, &doc_summary),
            ]),
        )
        .expect("fixture");

        run(
            input.clone(),
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inspect-metadata json");

        let text = fs::read_to_string(&output).expect("output");
        assert!(text.contains("\"name\": \"summary-information\""));
        assert!(text.contains("\"name\": \"title\""));
        assert!(text.contains("\"value\": \"Specimen\""));
        assert!(text.contains("\"name\": \"subject\""));
        assert!(text.contains("\"name\": \"author\""));
        assert!(text.contains("\"name\": \"last-saved-by\""));
        assert!(text.contains("\"name\": \"keywords\""));
        assert!(text.contains("\"name\": \"edit-time\""));
        assert!(text.contains("\"name\": \"last-printed\""));
        assert!(text.contains("\"name\": \"page-count\""));
        assert!(text.contains("\"name\": \"word-count\""));
        assert!(text.contains("\"name\": \"application-name\""));
        assert!(text.contains("\"name\": \"security\""));
        assert!(text.contains("\"name\": \"category\""));
        assert!(text.contains("\"name\": \"presentation-format\""));
        assert!(text.contains("\"name\": \"byte-count\""));
        assert!(text.contains("\"name\": \"manager\""));
        assert!(text.contains("\"name\": \"company\""));
        assert!(text.contains("\"name\": \"content-type\""));
        assert!(text.contains("\"name\": \"content-status\""));
        assert!(text.contains("\"name\": \"language\""));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn inspect_metadata_run_writes_text() {
        let input = test_support::temp_file("legacy_text", "doc");
        let output = test_support::temp_file("legacy_text", "txt");
        let summary = build_test_property_set_stream(&[
            (4, TestPropertyValue::Str("Analyst")),
            (3, TestPropertyValue::Str("Specimen subject")),
            (5, TestPropertyValue::Str("macro,ole")),
            (10, TestPropertyValue::I64(3600)),
            (11, TestPropertyValue::FileTime(10_000_000)),
            (12, TestPropertyValue::FileTime(123)),
            (14, TestPropertyValue::I32(7)),
            (15, TestPropertyValue::I32(321)),
            (18, TestPropertyValue::Str("Microsoft Excel")),
            (19, TestPropertyValue::U32(1)),
        ]);
        let doc_summary = build_test_property_set_stream(&[
            (2, TestPropertyValue::Str("Malware triage")),
            (3, TestPropertyValue::Str("Screen")),
            (4, TestPropertyValue::I32(2048)),
            (14, TestPropertyValue::WStr("Analyst")),
            (15, TestPropertyValue::WStr("ACME")),
            (26, TestPropertyValue::WStr("application/vnd.ms-excel")),
            (27, TestPropertyValue::WStr("final")),
            (28, TestPropertyValue::WStr("en-US")),
        ]);
        fs::write(
            &input,
            build_test_cfb(&[
                (SUMMARY_INFO_STREAM, &summary),
                (DOC_SUMMARY_INFO_STREAM, &doc_summary),
            ]),
        )
        .expect("fixture");

        run(
            input.clone(),
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inspect-metadata text");

        let text = fs::read_to_string(&output).expect("output");
        assert!(text.contains("Section: summary-information"));
        assert!(text.contains("subject: Specimen subject"));
        assert!(text.contains("author: Analyst"));
        assert!(text.contains("keywords: macro,ole"));
        assert!(text.contains("edit-time: 3600"));
        assert!(text.contains("last-printed: 1601-01-01T00:00:01Z"));
        assert!(text.contains("page-count: 7"));
        assert!(text.contains("word-count: 321"));
        assert!(text.contains("application-name: Microsoft Excel"));
        assert!(text.contains("security: 1"));
        assert!(text.contains("created: 1601-01-01T00:00:00Z"));
        assert!(text.contains("Raw: 123"));
        assert!(text.contains("category: Malware triage"));
        assert!(text.contains("presentation-format: Screen"));
        assert!(text.contains("byte-count: 2048"));
        assert!(text.contains("manager: Analyst"));
        assert!(text.contains("company: ACME"));
        assert!(text.contains("content-type: application/vnd.ms-excel"));
        assert!(text.contains("content-status: final"));
        assert!(text.contains("language: en-US"));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn format_metadata_text_renders_expected_fields() {
        let metadata = MetadataInspection {
            container: "cfb-ole".to_string(),
            section_count: 1,
            sections: vec![MetadataSection {
                name: "summary-information".to_string(),
                path: SUMMARY_INFO_STREAM.to_string(),
                property_count: 1,
                properties: vec![MetadataProperty {
                    id: 2,
                    name: "title".to_string(),
                    value_type: "lpstr".to_string(),
                    value: "Specimen".to_string(),
                    display_value: None,
                }],
            }],
        };
        let text = format_metadata_text(&metadata);
        assert!(text.contains("Container: cfb-ole"));
        assert!(text.contains("title: Specimen"));
        assert!(text.contains("Type: lpstr"));
    }
}
