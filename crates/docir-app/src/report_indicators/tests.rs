use super::IndicatorReport;
use crate::inspect_directory_bytes;
use crate::test_support::{
    build_test_cfb, patch_test_cfb_directory_entry, patch_test_cfb_header_u32,
    TestCfbDirectoryPatch,
};
use crate::ParsedDocument;
use docir_core::ir::{Document, IRNode};
use docir_core::security::{
    ActiveXControl, DdeField, DdeFieldType, ExternalRefType, ExternalReference, MacroModule,
    MacroModuleType, MacroProject, ThreatIndicator, ThreatIndicatorType, ThreatLevel,
};
use docir_core::types::{DocumentFormat, SourceSpan};
use docir_core::visitor::IrStore;

fn make_parsed_with_security() -> ParsedDocument {
    let mut store = IrStore::new();

    let mut project = MacroProject::new();
    project.name = Some("LegacyProject".to_string());
    project.container_path = Some("word/vbaProject.bin".to_string());
    project.storage_root = Some("VBA".to_string());
    project.has_auto_exec = true;
    project.auto_exec_procedures = vec!["AutoOpen".to_string()];
    project.is_protected = true;

    let mut module = MacroModule::new("Module1", MacroModuleType::Standard);
    module.stream_path = Some("VBA/Module1".to_string());
    module.procedures = vec!["AutoOpen".to_string()];
    project.modules.push(module.id);

    let mut ole = docir_core::security::OleObject::new();
    ole.source_path = Some("ObjectPool/1/Ole10Native".to_string());
    ole.embedded_payload_kind = Some("ole10native".to_string());
    ole.size_bytes = 128;

    let mut activex = ActiveXControl::new();
    activex.name = Some("DangerousControl".to_string());

    let external = ExternalReference::new(
        ExternalRefType::AttachedTemplate,
        "https://evil.test/template.dotm",
    );

    let dde = DdeField {
        field_type: DdeFieldType::DdeAuto,
        application: "cmd".to_string(),
        topic: None,
        item: Some("/c calc".to_string()),
        instruction: r#"DDEAUTO "cmd" "/c calc""#.to_string(),
        location: Some(SourceSpan {
            file_path: "word/document.xml".to_string(),
            relationship_id: None,
            xml_path: None,
            line: None,
            column: None,
        }),
    };

    let mut doc = Document::new(DocumentFormat::WordProcessing);
    doc.security.macro_project = Some(project.id);
    doc.security.ole_objects.push(ole.id);
    doc.security.activex_controls.push(activex.id);
    doc.security.external_refs.push(external.id);
    doc.security.dde_fields.push(dde);
    doc.security.threat_level = ThreatLevel::Critical;
    doc.security.threat_indicators.push(ThreatIndicator {
        indicator_type: ThreatIndicatorType::ExternalTemplate,
        severity: ThreatLevel::Medium,
        description: "Remote template relationship".to_string(),
        location: Some("word/_rels/document.xml.rels".to_string()),
        node_id: None,
    });
    let root_id = doc.id;

    store.insert(IRNode::MacroProject(project));
    store.insert(IRNode::MacroModule(module));
    store.insert(IRNode::OleObject(ole));
    store.insert(IRNode::ActiveXControl(activex));
    store.insert(IRNode::ExternalReference(external));
    store.insert(IRNode::Document(doc));

    ParsedDocument::new(docir_parser::parser::ParsedDocument {
        root_id,
        format: DocumentFormat::WordProcessing,
        store,
        metrics: None,
    })
}

#[test]
fn report_indicators_collects_expected_signals() {
    let parsed = make_parsed_with_security();
    let report = IndicatorReport::from_parsed(&parsed);
    assert_eq!(report.container, "zip-ooxml");
    assert_eq!(report.overall_risk, ThreatLevel::Critical);
    assert!(report.indicators.iter().any(|indicator| {
        indicator.key == "macros"
            && indicator.value == "1"
            && indicator.risk == ThreatLevel::Critical
    }));
    assert!(report.indicators.iter().any(|indicator| {
        indicator.key == "object-pool"
            && indicator
                .evidence
                .iter()
                .any(|value| value.contains("ObjectPool/1/Ole10Native"))
    }));
    assert!(report.indicators.iter().any(|indicator| {
        indicator.key == "suspicious-relationships" && indicator.value == "1"
    }));
}

#[test]
fn report_indicators_marks_absent_when_no_security_content_exists() {
    let mut store = IrStore::new();
    let doc = Document::new(DocumentFormat::WordProcessing);
    let root_id = doc.id;
    store.insert(IRNode::Document(doc));
    let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
        root_id,
        format: DocumentFormat::WordProcessing,
        store,
        metrics: None,
    });

    let report = IndicatorReport::from_parsed(&parsed);
    assert_eq!(report.overall_risk, ThreatLevel::None);
    assert!(report
        .indicators
        .iter()
        .all(|indicator| indicator.key == "format-container" || indicator.value == "absent"));
}

#[test]
fn report_indicators_includes_cfb_structural_anomalies_when_source_bytes_provided() {
    let mut store = IrStore::new();
    let doc = Document::new(DocumentFormat::WordProcessing);
    let root_id = doc.id;
    store.insert(IRNode::Document(doc));
    let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
        root_id,
        format: DocumentFormat::WordProcessing,
        store,
        metrics: None,
    });

    let cfb = build_test_cfb(&[("WordDocument", b"doc")]);
    let patched = patch_test_cfb_header_u32(&patch_test_cfb_header_u32(&cfb, 0x3C, 0), 0x40, 1);
    let report = IndicatorReport::from_parsed_with_bytes(&parsed, Some(&patched));
    assert!(report.indicators.iter().any(|indicator| {
        indicator.key == "cfb-structural-anomalies"
            && indicator.value != "absent"
            && indicator
                .evidence
                .iter()
                .any(|value| value.contains("structural-incoherence"))
    }));
    assert!(report.indicators.iter().any(|indicator| {
        indicator.key == "cfb-structural-score" && indicator.value == "medium"
    }));
    assert!(report
        .indicators
        .iter()
        .any(|indicator| { indicator.key == "cfb-directory-score" && indicator.value == "none" }));
    assert!(report
        .indicators
        .iter()
        .any(|indicator| { indicator.key == "cfb-sector-score" && indicator.value == "medium" }));
}

#[test]
fn report_indicators_surface_specific_structural_classes() {
    let mut store = IrStore::new();
    let doc = Document::new(DocumentFormat::WordProcessing);
    let root_id = doc.id;
    store.insert(IRNode::Document(doc));
    let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
        root_id,
        format: DocumentFormat::WordProcessing,
        store,
        metrics: None,
    });

    let base = build_test_cfb(&[
        ("WordDocument", b"doc"),
        ("VBA/PROJECT", b"meta"),
        ("VBA/Module1", b"code"),
        ("ObjectPool/1/Ole10Native", b"payload"),
    ]);
    let inspection = inspect_directory_bytes(&base).expect("directory");
    let word = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "WordDocument")
        .expect("word");
    let vba = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "VBA/PROJECT")
        .expect("vba");
    let objectpool = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "ObjectPool/1/Ole10Native")
        .expect("objectpool");

    let patched = patch_test_cfb_directory_entry(
        &patch_test_cfb_directory_entry(
            &patch_test_cfb_directory_entry(
                &base,
                vba.entry_index,
                TestCfbDirectoryPatch {
                    start_sector: Some(word.start_sector),
                    ..Default::default()
                },
            ),
            objectpool.entry_index,
            TestCfbDirectoryPatch {
                start_sector: Some(99),
                ..Default::default()
            },
        ),
        word.entry_index,
        TestCfbDirectoryPatch {
            start_sector: Some(98),
            ..Default::default()
        },
    );

    let report = IndicatorReport::from_parsed_with_bytes(&parsed, Some(&patched));
    assert!(report.indicators.iter().any(|indicator| {
        indicator.key == "cfb-objectpool-corruption" && indicator.value != "absent"
    }));
    assert!(report.indicators.iter().any(|indicator| {
        indicator.key == "cfb-vba-structure-anomalies" && indicator.value != "absent"
    }));
    assert!(report.indicators.iter().any(|indicator| {
        indicator.key == "cfb-main-stream-corruption" && indicator.value != "absent"
    }));
    assert!(report
        .indicators
        .iter()
        .any(|indicator| { indicator.key == "cfb-directory-score" && indicator.value != "none" }));
    assert!(report
        .indicators
        .iter()
        .any(|indicator| { indicator.key == "cfb-sector-score" && indicator.value != "none" }));
    assert!(report.indicators.iter().any(|indicator| {
        indicator.key == "cfb-dominant-anomaly-class"
            && matches!(indicator.value.as_str(), "shared-sector" | "invalid-start")
    }));
    assert!(report.indicators.iter().any(|indicator| {
        indicator.key == "cfb-objectpool-corruption"
            && indicator
                .evidence
                .iter()
                .any(|value| value.starts_with("objectpool:"))
    }));
    assert!(report.indicators.iter().any(|indicator| {
        indicator.key == "cfb-vba-structure-anomalies"
            && indicator
                .evidence
                .iter()
                .any(|value| value.starts_with("vba:"))
    }));
    assert!(report.indicators.iter().any(|indicator| {
        indicator.key == "cfb-main-stream-corruption"
            && indicator
                .evidence
                .iter()
                .any(|value| value.starts_with("main-stream:word:"))
    }));
}

#[test]
fn dominant_anomaly_class_uses_stable_tie_break_order() {
    let mut store = IrStore::new();
    let doc = Document::new(DocumentFormat::WordProcessing);
    let root_id = doc.id;
    store.insert(IRNode::Document(doc));
    let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
        root_id,
        format: DocumentFormat::WordProcessing,
        store,
        metrics: None,
    });
    let cfb = build_test_cfb(&[("WordDocument", b"doc")]);
    let patched = patch_test_cfb_header_u32(&patch_test_cfb_header_u32(&cfb, 0x3C, 0), 0x40, 1);
    let report = IndicatorReport::from_parsed_with_bytes(&parsed, Some(&patched));
    let dominant = report
        .indicators
        .iter()
        .find(|indicator| indicator.key == "cfb-dominant-anomaly-class")
        .expect("dominant anomaly indicator");
    assert_eq!(dominant.value, "mini-fat");
}

#[test]
fn structural_indicator_taxonomy_prefixes_remain_stable() {
    let mut store = IrStore::new();
    let doc = Document::new(DocumentFormat::WordProcessing);
    let root_id = doc.id;
    store.insert(IRNode::Document(doc));
    let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
        root_id,
        format: DocumentFormat::WordProcessing,
        store,
        metrics: None,
    });
    let base = build_test_cfb(&[
        ("WordDocument", b"doc"),
        ("VBA/PROJECT", b"vba"),
        ("ObjectPool/1/Ole10Native", b"obj"),
    ]);
    let inspection = inspect_directory_bytes(&base).expect("directory");
    let word = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "WordDocument")
        .expect("word");
    let vba = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "VBA/PROJECT")
        .expect("vba");
    let objectpool = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "ObjectPool/1/Ole10Native")
        .expect("objectpool");

    let patched = patch_test_cfb_directory_entry(
        &patch_test_cfb_directory_entry(
            &patch_test_cfb_directory_entry(
                &base,
                vba.entry_index,
                TestCfbDirectoryPatch {
                    start_sector: Some(word.start_sector),
                    ..Default::default()
                },
            ),
            objectpool.entry_index,
            TestCfbDirectoryPatch {
                start_sector: Some(99),
                ..Default::default()
            },
        ),
        word.entry_index,
        TestCfbDirectoryPatch {
            start_sector: Some(98),
            ..Default::default()
        },
    );
    let report = IndicatorReport::from_parsed_with_bytes(&parsed, Some(&patched));
    let evidence: Vec<&str> = report
        .indicators
        .iter()
        .flat_map(|indicator| indicator.evidence.iter().map(String::as_str))
        .collect();

    assert!(evidence.iter().any(|value| value.starts_with("directory:")));
    assert!(evidence.iter().any(|value| value.starts_with("sector:")));
    assert!(evidence.iter().any(|value| value.starts_with("health:")));
    assert!(evidence
        .iter()
        .any(|value| value.starts_with("objectpool:")));
    assert!(evidence.iter().any(|value| value.starts_with("vba:")));
    assert!(evidence
        .iter()
        .any(|value| value.starts_with("main-stream:")));
}
