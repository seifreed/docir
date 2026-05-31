use super::{
    DOC_SUMMARY_INFO_STREAM, MetadataSection, SUMMARY_INFO_STREAM, inspect_metadata_bytes,
};
use crate::test_support::{TestPropertyValue, build_test_cfb, build_test_property_set_stream};

#[test]
fn inspect_metadata_reads_summary_and_doc_summary_streams() {
    let summary = build_test_property_set_stream(&[
        (2, TestPropertyValue::Str("Specimen")),
        (4, TestPropertyValue::Str("Analyst")),
        (12, TestPropertyValue::FileTime(100)),
        (14, TestPropertyValue::I32(7)),
    ]);
    let doc_summary = build_test_property_set_stream(&[
        (15, TestPropertyValue::WStr("ACME")),
        (16, TestPropertyValue::Bool(true)),
    ]);
    let inspection = inspect_metadata_bytes(&build_test_cfb(&[
        (SUMMARY_INFO_STREAM, &summary),
        (DOC_SUMMARY_INFO_STREAM, &doc_summary),
    ]))
    .expect("inspection");

    assert_eq!(inspection.section_count, 2);
    let summary = inspection
        .sections
        .iter()
        .find(|section| section.name == "summary-information")
        .expect("summary section");
    assert!(summary.properties.iter().any(|prop| prop.name == "title"
        && prop.value == "Specimen"
        && prop.display_value.is_none()));
    assert!(
        summary
            .properties
            .iter()
            .any(|prop| prop.name == "page-count" && prop.value == "7")
    );
    assert!(summary.properties.iter().any(|prop| {
        prop.name == "created"
            && prop.value == "100"
            && prop
                .display_value
                .as_deref()
                .is_some_and(|value| value.ends_with('Z'))
    }));

    let doc_summary = inspection
        .sections
        .iter()
        .find(|section| section.name == "document-summary-information")
        .expect("doc summary section");
    assert!(
        doc_summary
            .properties
            .iter()
            .any(|prop| prop.name == "company" && prop.value == "ACME")
    );
    assert!(
        doc_summary
            .properties
            .iter()
            .any(|prop| prop.name == "links-dirty" && prop.value == "true")
    );
}

fn additional_scalar_property_streams() -> (Vec<u8>, Vec<u8>) {
    let summary = build_test_property_set_stream(&[
        (1, TestPropertyValue::U16(1200)),
        (3, TestPropertyValue::Str("Specimen subject")),
        (4, TestPropertyValue::Str("Analyst")),
        (5, TestPropertyValue::Str("macro,ole")),
        (6, TestPropertyValue::Str("sample comment")),
        (7, TestPropertyValue::Str("Normal.dotm")),
        (8, TestPropertyValue::Str("Responder")),
        (9, TestPropertyValue::Str("7")),
        (10, TestPropertyValue::I64(3600)),
        (14, TestPropertyValue::I16(3)),
        (15, TestPropertyValue::U32(42)),
        (18, TestPropertyValue::Str("Microsoft Excel")),
        (19, TestPropertyValue::U32(1)),
    ]);
    let doc_summary = build_test_property_set_stream(&[
        (2, TestPropertyValue::Str("Malware triage")),
        (4, TestPropertyValue::I64(2048)),
        (14, TestPropertyValue::WStr("Analyst")),
        (15, TestPropertyValue::WStr("ACME")),
        (29, TestPropertyValue::WStr("16.0")),
        (12, TestPropertyValue::WStr("Slides")),
        (13, TestPropertyValue::WStr("Part A")),
        (26, TestPropertyValue::WStr("application/vnd.ms-excel")),
        (27, TestPropertyValue::WStr("final")),
        (23, TestPropertyValue::F64(2.5)),
        (28, TestPropertyValue::WStr("en-US")),
    ]);

    (summary, doc_summary)
}

fn assert_metadata_properties(section: &MetadataSection, expected: &[(&str, &str, &str)]) {
    for (name, value_type, value) in expected {
        assert!(
            section.properties.iter().any(|prop| {
                prop.name == *name && prop.value_type == *value_type && prop.value == *value
            }),
            "missing property {name} with type {value_type} and value {value}"
        );
    }
}

fn assert_additional_summary_properties(summary: &MetadataSection) {
    assert_metadata_properties(
        summary,
        &[
            ("codepage", "u16", "1200"),
            ("subject", "lpstr", "Specimen subject"),
            ("author", "lpstr", "Analyst"),
            ("keywords", "lpstr", "macro,ole"),
            ("comments", "lpstr", "sample comment"),
            ("last-saved-by", "lpstr", "Responder"),
            ("template", "lpstr", "Normal.dotm"),
            ("revision-number", "lpstr", "7"),
            ("edit-time", "i64", "3600"),
            ("page-count", "i16", "3"),
            ("word-count", "u32", "42"),
            ("application-name", "lpstr", "Microsoft Excel"),
            ("security", "u32", "1"),
        ],
    );
}

fn assert_additional_doc_summary_properties(doc_summary: &MetadataSection) {
    assert_metadata_properties(
        doc_summary,
        &[
            ("category", "lpstr", "Malware triage"),
            ("byte-count", "i64", "2048"),
            ("heading-pairs", "lpwstr", "Slides"),
            ("titles-of-parts", "lpwstr", "Part A"),
            ("manager", "lpwstr", "Analyst"),
            ("company", "lpwstr", "ACME"),
            ("hlinks", "f64", "2.5"),
            ("content-type", "lpwstr", "application/vnd.ms-excel"),
            ("content-status", "lpwstr", "final"),
            ("language", "lpwstr", "en-US"),
            ("document-version", "lpwstr", "16.0"),
        ],
    );
}

#[test]
fn inspect_metadata_supports_additional_scalar_property_types() {
    let (summary, doc_summary) = additional_scalar_property_streams();
    let inspection = inspect_metadata_bytes(&build_test_cfb(&[
        (SUMMARY_INFO_STREAM, &summary),
        (DOC_SUMMARY_INFO_STREAM, &doc_summary),
    ]))
    .expect("inspection");

    let summary = inspection
        .sections
        .iter()
        .find(|section| section.name == "summary-information")
        .expect("summary section");
    assert_additional_summary_properties(summary);

    let doc_summary = inspection
        .sections
        .iter()
        .find(|section| section.name == "document-summary-information")
        .expect("doc summary section");
    assert_additional_doc_summary_properties(doc_summary);
}
