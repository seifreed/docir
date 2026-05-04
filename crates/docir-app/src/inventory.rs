use crate::ParsedDocument;
use docir_core::ir::{CustomXmlPart, Diagnostics, ExtensionPart, IRNode, MediaAsset};
use docir_core::security::{
    ActiveXControl, DdeField, ExternalReference, MacroModule, MacroProject, OleObject,
};
use docir_core::types::{DocumentFormat, NodeId};
use docir_parser::ole::{Cfb, CfbEntryType};
use serde::Serialize;

/// Container classification surfaced to inventory clients.
#[derive(Debug, Clone, Copy, Serialize, PartialEq, Eq)]
pub enum ContainerKind {
    ZipOoxml,
    ZipOdf,
    ZipHwpx,
    CfbOle,
    Rtf,
    Unknown,
}

/// High-level artifact kinds represented in the inventory.
#[derive(Debug, Clone, Copy, Serialize, PartialEq, Eq)]
pub enum InventoryArtifactKind {
    ContainerEntry,
    RelationshipGraph,
    ExtensionPart,
    MediaAsset,
    CustomXmlPart,
    MacroProject,
    MacroModule,
    OleObject,
    ActiveXControl,
    ExternalReference,
    DdeField,
    Diagnostics,
}

/// Per-document inventory summary.
#[derive(Debug, Clone, Serialize)]
pub struct ArtifactInventory {
    pub document_format: String,
    pub container_kind: ContainerKind,
    pub artifact_count: usize,
    pub macro_project_count: usize,
    pub macro_module_count: usize,
    pub ole_object_count: usize,
    pub activex_control_count: usize,
    pub external_reference_count: usize,
    pub dde_field_count: usize,
    pub artifacts: Vec<InventoryArtifact>,
}

/// A single inventory evidence item.
#[derive(Debug, Clone, Serialize)]
pub struct InventoryArtifact {
    pub kind: InventoryArtifactKind,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub node_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start_sector: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub created_filetime: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub modified_filetime: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub media_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sha256: Option<String>,
    pub details: String,
}

#[derive(Debug, Clone, Default)]
struct InventoryArtifactMeta {
    size_bytes: Option<u64>,
    start_sector: Option<u32>,
    created_filetime: Option<u64>,
    modified_filetime: Option<u64>,
    media_type: Option<String>,
    sha256: Option<String>,
}

impl ArtifactInventory {
    /// Builds a document-wide artifact inventory from the parsed IR.
    pub fn from_parsed(parsed: &ParsedDocument) -> Self {
        let mut inventory = Self {
            document_format: parsed.format().extension().to_string(),
            container_kind: classify_container(parsed),
            artifact_count: 0,
            macro_project_count: 0,
            macro_module_count: 0,
            ole_object_count: 0,
            activex_control_count: 0,
            external_reference_count: 0,
            dde_field_count: 0,
            artifacts: Vec::new(),
        };

        inventory.collect_ir_artifacts(parsed);

        if let Some(security) = parsed.security_info() {
            for dde in &security.dde_fields {
                inventory.push_dde_field(dde);
            }
        }

        inventory.artifact_count = inventory.artifacts.len();
        inventory
    }

    fn collect_ir_artifacts(&mut self, parsed: &ParsedDocument) {
        for (id, node) in parsed.store().iter() {
            match node {
                IRNode::RelationshipGraph(graph) => {
                    self.push_node_artifact(
                        InventoryArtifactKind::RelationshipGraph,
                        Some(*id),
                        Some(graph.source.clone()),
                        None,
                        InventoryArtifactMeta {
                            size_bytes: Some(graph.relationships.len() as u64),
                            ..InventoryArtifactMeta::default()
                        },
                        format!("{} relationship(s)", graph.relationships.len()),
                    );
                }
                IRNode::ExtensionPart(part) => self.push_extension_part(*id, part),
                IRNode::MediaAsset(asset) => self.push_media_asset(*id, asset),
                IRNode::CustomXmlPart(part) => self.push_custom_xml(*id, part),
                IRNode::MacroProject(project) => self.push_macro_project(*id, project),
                IRNode::MacroModule(module) => self.push_macro_module(*id, module),
                IRNode::OleObject(ole) => self.push_ole_object(*id, ole),
                IRNode::ActiveXControl(control) => self.push_activex(*id, control),
                IRNode::ExternalReference(ext) => self.push_external_reference(*id, ext),
                IRNode::Diagnostics(diag) => self.push_diagnostics(*id, diag),
                _ => {}
            }
        }
    }

    /// Builds a document-wide artifact inventory and enriches it with low-level CFB metadata.
    pub fn from_parsed_with_bytes(parsed: &ParsedDocument, input_bytes: &[u8]) -> Self {
        let mut inventory = Self::from_parsed(parsed);
        if inventory.container_kind == ContainerKind::CfbOle {
            inventory.append_cfb_container_entries(input_bytes);
            inventory.artifact_count = inventory.artifacts.len();
        }
        inventory
    }

    fn push_node_artifact(
        &mut self,
        kind: InventoryArtifactKind,
        node_id: Option<NodeId>,
        path: Option<String>,
        relationship_id: Option<String>,
        meta: InventoryArtifactMeta,
        details: String,
    ) {
        self.artifacts.push(InventoryArtifact {
            kind,
            node_id: node_id.map(|id| id.to_string()),
            path,
            relationship_id,
            size_bytes: meta.size_bytes,
            start_sector: meta.start_sector,
            created_filetime: meta.created_filetime,
            modified_filetime: meta.modified_filetime,
            media_type: meta.media_type,
            sha256: meta.sha256,
            details,
        });
    }

    fn push_extension_part(&mut self, id: NodeId, part: &ExtensionPart) {
        self.push_node_artifact(
            InventoryArtifactKind::ExtensionPart,
            Some(id),
            Some(part.path.clone()),
            None,
            InventoryArtifactMeta {
                size_bytes: Some(part.size_bytes),
                ..InventoryArtifactMeta::default()
            },
            format!("{:?}", part.kind),
        );
    }

    fn push_media_asset(&mut self, id: NodeId, asset: &MediaAsset) {
        self.push_node_artifact(
            InventoryArtifactKind::MediaAsset,
            Some(id),
            Some(asset.path.clone()),
            None,
            InventoryArtifactMeta {
                size_bytes: Some(asset.size_bytes),
                media_type: asset.content_type.clone(),
                sha256: asset.hash.clone(),
                ..InventoryArtifactMeta::default()
            },
            asset
                .content_type
                .clone()
                .unwrap_or_else(|| format!("{:?}", asset.media_type)),
        );
    }

    fn push_custom_xml(&mut self, id: NodeId, part: &CustomXmlPart) {
        self.push_node_artifact(
            InventoryArtifactKind::CustomXmlPart,
            Some(id),
            Some(part.path.clone()),
            None,
            InventoryArtifactMeta {
                size_bytes: Some(part.size_bytes),
                ..InventoryArtifactMeta::default()
            },
            part.root_element
                .clone()
                .unwrap_or_else(|| "customXml".to_string()),
        );
    }

    fn push_macro_project(&mut self, id: NodeId, project: &MacroProject) {
        self.macro_project_count += 1;
        self.push_node_artifact(
            InventoryArtifactKind::MacroProject,
            Some(id),
            project
                .container_path
                .clone()
                .or_else(|| project.span.as_ref().map(|s| s.file_path.clone())),
            None,
            InventoryArtifactMeta {
                size_bytes: Some(project.modules.len() as u64),
                ..InventoryArtifactMeta::default()
            },
            format_macro_project_details(project),
        );
    }

    fn push_macro_module(&mut self, id: NodeId, module: &MacroModule) {
        self.macro_module_count += 1;
        self.push_node_artifact(
            InventoryArtifactKind::MacroModule,
            Some(id),
            module
                .stream_path
                .clone()
                .or_else(|| module.span.as_ref().map(|s| s.file_path.clone())),
            None,
            InventoryArtifactMeta {
                size_bytes: module.decompressed_size.or(module.compressed_size),
                sha256: module.source_hash.clone(),
                ..InventoryArtifactMeta::default()
            },
            format_macro_module_details(module),
        );
    }

    fn push_ole_object(&mut self, id: NodeId, ole: &OleObject) {
        self.ole_object_count += 1;
        self.push_node_artifact(
            InventoryArtifactKind::OleObject,
            Some(id),
            ole.source_path
                .clone()
                .or_else(|| ole.span.as_ref().map(|s| s.file_path.clone())),
            ole.relationship_id.clone(),
            InventoryArtifactMeta {
                size_bytes: Some(ole.size_bytes),
                media_type: ole.mime_type.clone(),
                sha256: ole.data_hash.clone(),
                ..InventoryArtifactMeta::default()
            },
            ole.embedded_file_name
                .clone()
                .or_else(|| ole.class_name.clone())
                .or_else(|| ole.prog_id.clone())
                .unwrap_or_else(|| "ole object".to_string()),
        );
    }

    fn push_activex(&mut self, id: NodeId, control: &ActiveXControl) {
        self.activex_control_count += 1;
        self.push_node_artifact(
            InventoryArtifactKind::ActiveXControl,
            Some(id),
            control.span.as_ref().map(|s| s.file_path.clone()),
            control
                .span
                .as_ref()
                .and_then(|s| s.relationship_id.clone()),
            InventoryArtifactMeta::default(),
            control
                .name
                .clone()
                .or_else(|| control.prog_id.clone())
                .unwrap_or_else(|| "activex".to_string()),
        );
    }

    fn push_external_reference(&mut self, id: NodeId, ext: &ExternalReference) {
        self.external_reference_count += 1;
        self.push_node_artifact(
            InventoryArtifactKind::ExternalReference,
            Some(id),
            ext.span.as_ref().map(|s| s.file_path.clone()),
            ext.relationship_id.clone(),
            InventoryArtifactMeta::default(),
            format!("{:?}: {}", ext.ref_type, ext.target),
        );
    }

    fn push_dde_field(&mut self, dde: &DdeField) {
        self.dde_field_count += 1;
        self.push_node_artifact(
            InventoryArtifactKind::DdeField,
            None,
            dde.location.as_ref().map(|s| s.file_path.clone()),
            dde.location
                .as_ref()
                .and_then(|s| s.relationship_id.clone()),
            InventoryArtifactMeta::default(),
            dde.instruction.clone(),
        );
    }

    fn push_diagnostics(&mut self, id: NodeId, diag: &Diagnostics) {
        self.push_node_artifact(
            InventoryArtifactKind::Diagnostics,
            Some(id),
            diag.span.as_ref().map(|s| s.file_path.clone()),
            None,
            InventoryArtifactMeta {
                size_bytes: Some(diag.entries.len() as u64),
                ..InventoryArtifactMeta::default()
            },
            format!("{} diagnostic entrie(s)", diag.entries.len()),
        );
    }

    fn append_cfb_container_entries(&mut self, input_bytes: &[u8]) {
        let Ok(cfb) = Cfb::parse(input_bytes.to_vec()) else {
            return;
        };

        for entry in cfb.list_entries() {
            self.push_node_artifact(
                InventoryArtifactKind::ContainerEntry,
                None,
                Some(entry.path.clone()),
                None,
                InventoryArtifactMeta {
                    size_bytes: Some(entry.size),
                    start_sector: Some(entry.start_sector),
                    created_filetime: entry.created_filetime,
                    modified_filetime: entry.modified_filetime,
                    media_type: None,
                    sha256: None,
                },
                classify_cfb_entry_detail(&entry.path, entry.entry_type),
            );
        }
    }
}

fn classify_container(parsed: &ParsedDocument) -> ContainerKind {
    if parsed
        .document()
        .and_then(|doc| doc.span.as_ref())
        .map(|span| span.file_path.starts_with("cfb:/"))
        .unwrap_or(false)
    {
        return ContainerKind::CfbOle;
    }

    match parsed.format() {
        DocumentFormat::WordProcessing
        | DocumentFormat::Spreadsheet
        | DocumentFormat::Presentation => ContainerKind::ZipOoxml,
        DocumentFormat::OdfText
        | DocumentFormat::OdfSpreadsheet
        | DocumentFormat::OdfPresentation => ContainerKind::ZipOdf,
        DocumentFormat::Hwpx => ContainerKind::ZipHwpx,
        DocumentFormat::Hwp => ContainerKind::CfbOle,
        DocumentFormat::Rtf => ContainerKind::Rtf,
    }
}

fn classify_cfb_entry_detail(path: &str, entry_type: CfbEntryType) -> String {
    match entry_type {
        CfbEntryType::RootStorage => "root-storage".to_string(),
        CfbEntryType::Storage => classify_cfb_storage(path),
        CfbEntryType::Stream => classify_cfb_stream(path),
    }
}

fn classify_cfb_storage(path: &str) -> String {
    let upper = path.to_ascii_uppercase();
    if upper == "OBJECTPOOL" || upper.starts_with("OBJECTPOOL/") {
        "embedded-object-storage".to_string()
    } else if upper == "VBA" || upper.ends_with("/VBA") {
        "vba-storage".to_string()
    } else {
        "storage".to_string()
    }
}

fn classify_cfb_stream(path: &str) -> String {
    let upper = path.to_ascii_uppercase();
    if upper == "WORDDOCUMENT" {
        "word-main-stream".to_string()
    } else if upper == "WORKBOOK" || upper == "BOOK" {
        "excel-main-stream".to_string()
    } else if upper == "POWERPOINT DOCUMENT" {
        "powerpoint-main-stream".to_string()
    } else if upper.ends_with("/PROJECT") || upper == "PROJECT" {
        "vba-project-metadata".to_string()
    } else if upper.contains("/VBA/") || upper.starts_with("VBA/") {
        "vba-module-stream".to_string()
    } else if upper.ends_with("OLE10NATIVE") {
        "ole-native-payload".to_string()
    } else if upper.ends_with("/PACKAGE") {
        "package-payload".to_string()
    } else if upper.ends_with("/CONTENTS") {
        "embedded-contents".to_string()
    } else {
        "stream".to_string()
    }
}

fn format_macro_project_details(project: &MacroProject) -> String {
    let mut parts = vec![
        format!(
            "name={}",
            project
                .name
                .clone()
                .unwrap_or_else(|| "macro project".to_string())
        ),
        format!("protected={}", project.is_protected),
        format!("modules={}", project.modules.len()),
        format!("references={}", project.references.len()),
    ];

    if !project.auto_exec_procedures.is_empty() {
        parts.push(format!(
            "autoexec={}",
            project.auto_exec_procedures.join(",")
        ));
    }
    if project.digital_signature.is_some() {
        parts.push("signed=true".to_string());
    }

    parts.join("; ")
}

fn format_macro_module_details(module: &MacroModule) -> String {
    let mut parts = vec![
        format!("type={}", lowercase_debug(module.module_type)),
        format!("state={}", lowercase_debug(module.extraction_state)),
    ];

    if !module.procedures.is_empty() {
        parts.push(format!("procedures={}", module.procedures.join(",")));
    }
    if !module.extraction_errors.is_empty() {
        parts.push(format!("errors={}", module.extraction_errors.join(" | ")));
    }

    parts.join("; ")
}

fn lowercase_debug<T: std::fmt::Debug>(value: T) -> String {
    format!("{:?}", value).to_ascii_lowercase()
}

#[cfg(test)]
mod tests {
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
}
