use super::*;
use crate::test_support::build_test_cfb;
use docir_core::ir::{Document, IRNode};
use docir_core::security::{MacroExtractionState, MacroModuleType};
use docir_core::types::SourceSpan;
use docir_core::visitor::IrStore;

#[test]
fn inventory_collects_macro_ole_and_dde_evidence() {
    let mut store = IrStore::new();
    let mut doc = Document::new(DocumentFormat::WordProcessing);

    let mut project = MacroProject::new();
    project.container_path = Some("word/vbaProject.bin".to_string());
    let project_id = project.id;

    let mut module = MacroModule::new("Module1", MacroModuleType::Standard);
    module.stream_path = Some("VBA/Module1".to_string());
    module.extraction_state = MacroExtractionState::Extracted;
    module.decompressed_size = Some(12);
    let module_id = module.id;
    project.modules.push(module_id);

    let mut ole = OleObject::new();
    ole.source_path = Some("word/embeddings/object1.bin".to_string());
    ole.size_bytes = 42;
    let ole_id = ole.id;

    doc.security.macro_project = Some(project_id);
    doc.security.ole_objects.push(ole_id);
    doc.security.dde_fields.push(DdeField {
        field_type: docir_core::security::DdeFieldType::DdeAuto,
        application: "cmd".to_string(),
        topic: None,
        item: None,
        instruction: r#"DDEAUTO "cmd" "/c calc""#.to_string(),
        location: None,
    });

    let root_id = doc.id;
    store.insert(IRNode::MacroProject(project));
    store.insert(IRNode::MacroModule(module));
    store.insert(IRNode::OleObject(ole));
    store.insert(IRNode::Document(doc));

    let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
        root_id,
        format: DocumentFormat::WordProcessing,
        store,
        metrics: None,
    });

    let inventory = ArtifactInventory::from_parsed(&parsed);
    assert_eq!(inventory.container_kind, ContainerKind::ZipOoxml);
    assert_eq!(inventory.macro_project_count, 1);
    assert_eq!(inventory.macro_module_count, 1);
    assert_eq!(inventory.ole_object_count, 1);
    assert_eq!(inventory.dde_field_count, 1);
    assert!(inventory.artifact_count >= 4);
    let project_artifact = inventory
        .artifacts
        .iter()
        .find(|artifact| artifact.kind == InventoryArtifactKind::MacroProject)
        .expect("macro project");
    assert!(project_artifact.details.contains("name=macro project"));
    assert!(project_artifact.details.contains("protected=false"));
    let module_artifact = inventory
        .artifacts
        .iter()
        .find(|artifact| artifact.kind == InventoryArtifactKind::MacroModule)
        .expect("macro module");
    assert!(module_artifact.details.contains("type=standard"));
    assert!(module_artifact.details.contains("state=extracted"));
}

#[test]
fn inventory_marks_legacy_office_document_as_cfb() {
    let mut store = IrStore::new();
    let mut doc = Document::new(DocumentFormat::WordProcessing);
    doc.span = Some(docir_core::types::SourceSpan::new("cfb:/"));
    let root_id = doc.id;
    store.insert(IRNode::Document(doc));

    let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
        root_id,
        format: DocumentFormat::WordProcessing,
        store,
        metrics: None,
    });

    let inventory = ArtifactInventory::from_parsed(&parsed);
    assert_eq!(inventory.container_kind, ContainerKind::CfbOle);
}

#[test]
fn inventory_with_bytes_includes_cfb_container_entries() {
    let mut store = IrStore::new();
    let mut doc = Document::new(DocumentFormat::WordProcessing);
    doc.span = Some(SourceSpan::new("cfb:/"));
    let root_id = doc.id;
    store.insert(IRNode::Document(doc));

    let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
        root_id,
        format: DocumentFormat::WordProcessing,
        store,
        metrics: None,
    });

    let bytes = build_test_cfb(&[("WordDocument", b"main")]);
    let inventory = ArtifactInventory::from_parsed_with_bytes(&parsed, &bytes);
    assert!(inventory.artifacts.iter().any(|artifact| {
        artifact.kind == InventoryArtifactKind::ContainerEntry
            && artifact.path.as_deref() == Some("Root Entry")
    }));
}
