//! Build an artifact inventory for a document.

use anyhow::Result;
use docir_app::{ArtifactInventory, ParserConfig};
use serde::Serialize;
use std::fs;
use std::path::PathBuf;

use crate::commands::util::{
    build_app, push_bullet_line, push_labeled_line, write_json_output, write_text_output,
};

#[derive(Debug, Serialize)]
struct InventoryResult {
    inventory: ArtifactInventory,
}

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let app = build_app(parser_config);
    let source_bytes = fs::read(&input)?;
    let parsed = app.parse_bytes(&source_bytes)?;
    let inventory = app.build_inventory_with_bytes(&parsed, &source_bytes);

    if json {
        return write_json_output(&InventoryResult { inventory }, pretty, output);
    }

    let text = format_inventory_text(&inventory);
    write_text_output(&text, output)
}

fn format_inventory_text(inventory: &ArtifactInventory) -> String {
    let mut out = String::new();
    push_labeled_line(&mut out, 0, "Format", &inventory.document_format);
    push_labeled_line(
        &mut out,
        0,
        "Container",
        format!("{:?}", inventory.container_kind),
    );
    push_labeled_line(&mut out, 0, "Artifacts", inventory.artifact_count);
    push_labeled_line(&mut out, 0, "Macro Projects", inventory.macro_project_count);
    push_labeled_line(&mut out, 0, "Macro Modules", inventory.macro_module_count);
    push_labeled_line(&mut out, 0, "OLE Objects", inventory.ole_object_count);
    push_labeled_line(
        &mut out,
        0,
        "ActiveX Controls",
        inventory.activex_control_count,
    );
    push_labeled_line(
        &mut out,
        0,
        "External References",
        inventory.external_reference_count,
    );
    push_labeled_line(&mut out, 0, "DDE Fields", inventory.dde_field_count);

    if !inventory.artifacts.is_empty() {
        out.push_str("\nEvidence:\n");
        for artifact in &inventory.artifacts {
            push_bullet_line(
                &mut out,
                2,
                inventory_kind_label(artifact.kind),
                &artifact.details,
            );
            if let Some(path) = artifact.path.as_deref() {
                push_labeled_line(&mut out, 4, "Path", path);
            }
            if let Some(node_id) = artifact.node_id.as_deref() {
                push_labeled_line(&mut out, 4, "Node", node_id);
            }
            if let Some(relationship_id) = artifact.relationship_id.as_deref() {
                push_labeled_line(&mut out, 4, "Relationship", relationship_id);
            }
            if let Some(size_bytes) = artifact.size_bytes {
                push_labeled_line(&mut out, 4, "Size", format!("{size_bytes} bytes"));
            }
            if let Some(media_type) = artifact.media_type.as_deref() {
                push_labeled_line(&mut out, 4, "Media Type", media_type);
            }
            if let Some(sha256) = artifact.sha256.as_deref() {
                push_labeled_line(&mut out, 4, "SHA-256", sha256);
            }
            if let Some(start_sector) = artifact.start_sector {
                push_labeled_line(&mut out, 4, "CFB Sector", start_sector);
            }
            if let Some(created) = artifact.created_filetime {
                push_labeled_line(&mut out, 4, "CFB Created", created);
            }
            if let Some(modified) = artifact.modified_filetime {
                push_labeled_line(&mut out, 4, "CFB Modified", modified);
            }
        }
    }

    out
}

fn inventory_kind_label(kind: docir_app::InventoryArtifactKind) -> &'static str {
    match kind {
        docir_app::InventoryArtifactKind::ContainerEntry => "Container Entry",
        docir_app::InventoryArtifactKind::RelationshipGraph => "Relationship Graph",
        docir_app::InventoryArtifactKind::ExtensionPart => "Extension Part",
        docir_app::InventoryArtifactKind::MediaAsset => "Media Asset",
        docir_app::InventoryArtifactKind::CustomXmlPart => "Custom XML Part",
        docir_app::InventoryArtifactKind::MacroProject => "Macro Project",
        docir_app::InventoryArtifactKind::MacroModule => "Macro Module",
        docir_app::InventoryArtifactKind::OleObject => "OLE Object",
        docir_app::InventoryArtifactKind::ActiveXControl => "ActiveX Control",
        docir_app::InventoryArtifactKind::ExternalReference => "External Reference",
        docir_app::InventoryArtifactKind::DdeField => "DDE Field",
        docir_app::InventoryArtifactKind::Diagnostics => "Diagnostics",
    }
}

#[cfg(test)]
mod tests {
    use super::{format_inventory_text, run};
    use docir_app::{
        test_support::build_test_cfb, ArtifactInventory, ContainerKind, InventoryArtifact,
        InventoryArtifactKind, ParserConfig,
    };
    use std::fs;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn fixture(name: &str) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../fixtures/ooxml")
            .join(name)
    }

    fn temp_file(name: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_inventory_{name}_{nanos}.json"))
    }

    fn temp_input(name: &str, ext: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_inventory_{name}_{nanos}.{ext}"))
    }

    #[test]
    fn inventory_run_writes_json() {
        let output = temp_file("json");
        run(
            fixture("minimal.docx"),
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inventory json");
        let text = fs::read_to_string(&output).expect("inventory output");
        assert!(text.contains("\"inventory\""));
        let _ = fs::remove_file(output);
    }

    #[test]
    fn inventory_run_writes_text() {
        let output = temp_file("text");
        run(
            fixture("minimal.docx"),
            false,
            false,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inventory text");
        let text = fs::read_to_string(&output).expect("inventory output");
        assert!(text.contains("Format:"));
        assert!(text.contains("Evidence:"));
        let _ = fs::remove_file(output);
    }

    #[test]
    fn inventory_run_reports_vba_evidence_for_legacy_doc() {
        let input = temp_input("legacy_vba", "doc");
        let output = temp_file("legacy_vba");
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="LegacyProject"
Module=Module1/Module1
DPB="AAAA"
Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation
"#,
            ),
            ("VBA/Module1", b"Sub AutoOpen()\nEnd Sub\n"),
        ]);
        fs::write(&input, bytes).expect("fixture");

        run(
            input.clone(),
            false,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("inventory text");

        let text = fs::read_to_string(&output).expect("inventory output");
        assert!(text.contains("Macro Projects: 1"));
        assert!(text.contains("Macro Modules: 1"));
        assert!(text.contains("name=LegacyProject"));
        assert!(text.contains("protected=true"));
        assert!(text.contains("autoexec=AutoOpen"));
        assert!(text.contains("type=standard"));
        assert!(text.contains("procedures=AutoOpen"));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn format_inventory_text_includes_vba_and_ole_details() {
        let inventory = ArtifactInventory {
            document_format: "doc".to_string(),
            container_kind: ContainerKind::CfbOle,
            artifact_count: 3,
            macro_project_count: 1,
            macro_module_count: 1,
            ole_object_count: 1,
            activex_control_count: 0,
            external_reference_count: 0,
            dde_field_count: 0,
            artifacts: vec![
                InventoryArtifact {
                    kind: InventoryArtifactKind::MacroProject,
                    node_id: Some("node_1".to_string()),
                    path: Some("VBA".to_string()),
                    relationship_id: None,
                    size_bytes: Some(2),
                    start_sector: None,
                    created_filetime: None,
                    modified_filetime: None,
                    media_type: None,
                    sha256: None,
                    details: "MixedProtectedPartial".to_string(),
                },
                InventoryArtifact {
                    kind: InventoryArtifactKind::MacroModule,
                    node_id: Some("node_2".to_string()),
                    path: Some("VBA/Helper".to_string()),
                    relationship_id: None,
                    size_bytes: None,
                    start_sector: None,
                    created_filetime: None,
                    modified_filetime: None,
                    media_type: None,
                    sha256: Some("a".repeat(64)),
                    details: "class".to_string(),
                },
                InventoryArtifact {
                    kind: InventoryArtifactKind::OleObject,
                    node_id: Some("node_3".to_string()),
                    path: Some("ObjectPool/1/Ole10Native".to_string()),
                    relationship_id: Some("rId7".to_string()),
                    size_bytes: Some(128),
                    start_sector: Some(7),
                    created_filetime: Some(123),
                    modified_filetime: Some(456),
                    media_type: Some("application/x-dosexec".to_string()),
                    sha256: Some("b".repeat(64)),
                    details: "ole-native-payload".to_string(),
                },
            ],
        };

        let text = format_inventory_text(&inventory);
        assert!(text.contains("Format: doc"));
        assert!(text.contains("Container: CfbOle"));
        assert!(text.contains("Macro Projects: 1"));
        assert!(text.contains("OLE Objects: 1"));
        assert!(text.contains("- Macro Project: MixedProtectedPartial"));
        assert!(text.contains("Path: VBA"));
        assert!(text.contains("Node: node_1"));
        assert!(text.contains("- Macro Module: class"));
        assert!(text.contains("Path: VBA/Helper"));
        assert!(text.contains(&format!("SHA-256: {}", "a".repeat(64))));
        assert!(text.contains("- OLE Object: ole-native-payload"));
        assert!(text.contains("Relationship: rId7"));
        assert!(text.contains("Size: 128 bytes"));
        assert!(text.contains("Media Type: application/x-dosexec"));
        assert!(text.contains(&format!("SHA-256: {}", "b".repeat(64))));
        assert!(text.contains("CFB Sector: 7"));
        assert!(text.contains("CFB Created: 123"));
        assert!(text.contains("CFB Modified: 456"));
    }
}
