use super::*;
use crate::{ArtifactInventory, ContainerKind, InventoryArtifact, InventoryArtifactKind};

#[test]
fn phase0_manifest_from_inventory_maps_cfb_metadata() {
    let inventory = ArtifactInventory {
        document_format: "doc".to_string(),
        container_kind: ContainerKind::CfbOle,
        artifact_count: 1,
        macro_project_count: 0,
        macro_module_count: 0,
        ole_object_count: 0,
        activex_control_count: 0,
        external_reference_count: 0,
        dde_field_count: 0,
        artifacts: vec![InventoryArtifact {
            kind: InventoryArtifactKind::ContainerEntry,
            node_id: None,
            path: Some("ObjectPool/1/Ole10Native".to_string()),
            relationship_id: None,
            size_bytes: Some(128),
            start_sector: Some(7),
            created_filetime: Some(123),
            modified_filetime: Some(456),
            media_type: None,
            sha256: None,
            details: "ole-native-payload".to_string(),
        }],
    };

    let export = Phase0ArtifactManifestExport::from_inventory(
        &inventory,
        ExportDocumentRef::new("doc", "doc", Some("/tmp/sample.doc".to_string())),
    );

    assert_eq!(export.artifacts[0].kind, "container-entry");
    assert_eq!(export.artifacts[0].locator.offset, Some(7));
    assert_eq!(
        export.artifacts[0].locator.container.as_deref(),
        Some("/tmp/sample.doc")
    );
    assert_eq!(export.artifacts[0].diagnostics.len(), 2);
}

#[test]
fn phase0_manifest_from_inventory_preserves_vba_role_details() {
    let inventory = ArtifactInventory {
        document_format: "doc".to_string(),
        container_kind: ContainerKind::CfbOle,
        artifact_count: 2,
        macro_project_count: 1,
        macro_module_count: 1,
        ole_object_count: 0,
        activex_control_count: 0,
        external_reference_count: 0,
        dde_field_count: 0,
        artifacts: vec![
            InventoryArtifact {
                kind: InventoryArtifactKind::MacroProject,
                node_id: Some("node_1".to_string()),
                path: Some("VBA/PROJECT".to_string()),
                relationship_id: None,
                size_bytes: Some(2),
                start_sector: None,
                created_filetime: None,
                modified_filetime: None,
                media_type: None,
                sha256: None,
                details:
                    "name=LegacyProject; protected=true; modules=2; references=1; autoexec=AutoOpen"
                        .to_string(),
            },
            InventoryArtifact {
                kind: InventoryArtifactKind::MacroModule,
                node_id: Some("node_2".to_string()),
                path: Some("VBA/Module1".to_string()),
                relationship_id: None,
                size_bytes: Some(24),
                start_sector: None,
                created_filetime: None,
                modified_filetime: None,
                media_type: None,
                sha256: None,
                details:
                    "type=standard; state=extracted; procedures=AutoOpen; errors=Missing stream VBA/MissingMod"
                        .to_string(),
            },
        ],
    };

    let export = Phase0ArtifactManifestExport::from_inventory(
        &inventory,
        ExportDocumentRef::new("doc", "doc", Some("/tmp/sample.doc".to_string())),
    );

    assert_eq!(export.artifacts[0].kind, "vba-project");
    assert_eq!(
        export.artifacts[0].role.as_deref(),
        Some("name=LegacyProject; protected=true; modules=2; references=1; autoexec=AutoOpen")
    );
    assert_eq!(export.artifacts[1].kind, "vba-module-source");
    assert_eq!(
        export.artifacts[1].role.as_deref(),
        Some(
            "type=standard; state=extracted; procedures=AutoOpen; errors=Missing stream VBA/MissingMod"
        )
    );
}

#[test]
fn phase0_manifest_from_inventory_preserves_media_type() {
    let inventory = ArtifactInventory {
        document_format: "docx".to_string(),
        container_kind: ContainerKind::ZipOoxml,
        artifact_count: 1,
        macro_project_count: 0,
        macro_module_count: 0,
        ole_object_count: 0,
        activex_control_count: 0,
        external_reference_count: 0,
        dde_field_count: 0,
        artifacts: vec![InventoryArtifact {
            kind: InventoryArtifactKind::MediaAsset,
            node_id: Some("node_1".to_string()),
            path: Some("word/media/image1.png".to_string()),
            relationship_id: None,
            size_bytes: Some(128),
            start_sector: None,
            created_filetime: None,
            modified_filetime: None,
            media_type: Some("image/png".to_string()),
            sha256: None,
            details: "image/png".to_string(),
        }],
    };

    let export = Phase0ArtifactManifestExport::from_inventory(
        &inventory,
        ExportDocumentRef::new("docx", "docx", Some("/tmp/sample.docx".to_string())),
    );

    assert_eq!(export.artifacts[0].kind, "embedded-file");
    assert_eq!(export.artifacts[0].media_type.as_deref(), Some("image/png"));
}

#[test]
fn phase0_manifest_from_inventory_preserves_sha256() {
    let inventory = ArtifactInventory {
        document_format: "doc".to_string(),
        container_kind: ContainerKind::CfbOle,
        artifact_count: 2,
        macro_project_count: 0,
        macro_module_count: 1,
        ole_object_count: 1,
        activex_control_count: 0,
        external_reference_count: 0,
        dde_field_count: 0,
        artifacts: vec![
            InventoryArtifact {
                kind: InventoryArtifactKind::MacroModule,
                node_id: Some("node_1".to_string()),
                path: Some("VBA/Module1".to_string()),
                relationship_id: None,
                size_bytes: Some(24),
                start_sector: None,
                created_filetime: None,
                modified_filetime: None,
                media_type: None,
                sha256: Some("a".repeat(64)),
                details: "type=standard; state=extracted".to_string(),
            },
            InventoryArtifact {
                kind: InventoryArtifactKind::OleObject,
                node_id: Some("node_2".to_string()),
                path: Some("ObjectPool/1/Ole10Native".to_string()),
                relationship_id: None,
                size_bytes: Some(128),
                start_sector: Some(7),
                created_filetime: None,
                modified_filetime: None,
                media_type: Some("application/x-dosexec".to_string()),
                sha256: Some("b".repeat(64)),
                details: "dropper.exe".to_string(),
            },
        ],
    };

    let export = Phase0ArtifactManifestExport::from_inventory(
        &inventory,
        ExportDocumentRef::new("doc", "doc", Some("/tmp/sample.doc".to_string())),
    );

    let hash_a = "a".repeat(64);
    let hash_b = "b".repeat(64);
    assert_eq!(export.artifacts[0].sha256.as_deref(), Some(hash_a.as_str()));
    assert_eq!(export.artifacts[1].sha256.as_deref(), Some(hash_b.as_str()));
}

#[test]
fn phase0_manifest_from_extraction_maps_legacy_cfb_metadata() {
    let mut manifest = ExtractionManifest::new();
    manifest.source_document = Some("/tmp/sample.doc".to_string());
    let mut artifact = docir_core::ExtractedArtifact::new(
        "legacy-ole-object-1",
        docir_core::ExtractedArtifactKind::OleObject,
    );
    artifact.source_path = Some("ObjectPool/1/Ole10Native".to_string());
    artifact.size_bytes = Some(64);
    artifact.start_sector = Some(9);
    artifact.created_filetime = Some(111);
    artifact.modified_filetime = Some(222);
    manifest.artifacts.push(artifact);

    let export = Phase0ArtifactManifestExport::from_manifest(
        &manifest,
        ExportDocumentRef::new("doc", "doc", Some("/tmp/sample.doc".to_string())),
    );

    assert_eq!(export.artifacts[0].locator.offset, Some(9));
    assert_eq!(export.artifacts[0].diagnostics.len(), 2);
}
