use crate::ParsedDocument;
use docir_core::ir::IRNode;
use docir_core::security::{
    MacroExtractionState, MacroModule, MacroModuleType, MacroProject, SuspiciousCall,
};
use serde::Serialize;

/// High-level recognition status for a VBA extraction attempt.
#[derive(Debug, Clone, Copy, Serialize, PartialEq, Eq)]
pub enum VbaRecognitionStatus {
    Absent,
    Recognized,
    Extracted,
    Partial,
    Error,
}

/// Structured VBA recognition report for downstream tools.
#[derive(Debug, Clone, Serialize)]
pub struct VbaRecognitionReport {
    pub document_format: String,
    pub status: VbaRecognitionStatus,
    pub projects: Vec<VbaProjectReport>,
}

/// Report for a single VBA project.
#[derive(Debug, Clone, Serialize)]
pub struct VbaProjectReport {
    pub node_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub project_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub container_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub storage_root: Option<String>,
    pub is_protected: bool,
    pub has_auto_exec: bool,
    #[serde(default)]
    pub auto_exec_procedures: Vec<String>,
    #[serde(default)]
    pub references: Vec<String>,
    pub status: VbaRecognitionStatus,
    pub modules: Vec<VbaModuleReport>,
}

/// Report for a single VBA module.
#[derive(Debug, Clone, Serialize)]
pub struct VbaModuleReport {
    pub node_id: String,
    pub name: String,
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stream_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stream_path: Option<String>,
    pub status: VbaRecognitionStatus,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_encoding: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_hash: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub compressed_size: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub decompressed_size: Option<u64>,
    #[serde(default)]
    pub procedures: Vec<String>,
    #[serde(default)]
    pub suspicious_calls: Vec<String>,
    #[serde(default)]
    pub extraction_errors: Vec<String>,
}

impl VbaRecognitionReport {
    /// Builds a recognition report from the parsed IR.
    pub fn from_parsed(parsed: &ParsedDocument, include_source: bool) -> Self {
        let mut projects = Vec::new();

        for node in parsed.store().values() {
            let IRNode::MacroProject(project) = node else {
                continue;
            };
            projects.push(VbaProjectReport::from_project(
                parsed,
                project,
                include_source,
            ));
        }

        let status = aggregate_project_status(&projects);
        Self {
            document_format: parsed.format().extension().to_string(),
            status,
            projects,
        }
    }
}

impl VbaProjectReport {
    fn from_project(parsed: &ParsedDocument, project: &MacroProject, include_source: bool) -> Self {
        let modules: Vec<_> = project
            .modules
            .iter()
            .filter_map(|id| match parsed.store().get(*id) {
                Some(IRNode::MacroModule(module)) => {
                    Some(VbaModuleReport::from_module(module, include_source))
                }
                _ => None,
            })
            .collect();

        let status = aggregate_module_status(&modules);
        Self {
            node_id: project.id.to_string(),
            project_name: project.name.clone(),
            container_path: project.container_path.clone(),
            storage_root: project.storage_root.clone(),
            is_protected: project.is_protected,
            has_auto_exec: project.has_auto_exec,
            auto_exec_procedures: project.auto_exec_procedures.clone(),
            references: project.references.iter().map(|r| r.name.clone()).collect(),
            status,
            modules,
        }
    }
}

impl VbaModuleReport {
    fn from_module(module: &MacroModule, include_source: bool) -> Self {
        let source_text = include_source.then(|| module.source_code.clone()).flatten();
        Self {
            node_id: module.id.to_string(),
            name: module.name.clone(),
            kind: module_kind_name(module.module_type).to_string(),
            stream_name: module.stream_name.clone(),
            stream_path: module.stream_path.clone(),
            status: map_module_state(
                module.extraction_state,
                &module.extraction_errors,
                include_source,
            ),
            source_text,
            source_encoding: module.source_encoding.clone(),
            source_hash: module.source_hash.clone(),
            compressed_size: module.compressed_size,
            decompressed_size: module.decompressed_size,
            procedures: module.procedures.clone(),
            suspicious_calls: module
                .suspicious_calls
                .iter()
                .map(format_suspicious_call)
                .collect(),
            extraction_errors: module.extraction_errors.clone(),
        }
    }
}

fn aggregate_project_status(projects: &[VbaProjectReport]) -> VbaRecognitionStatus {
    if projects.is_empty() {
        return VbaRecognitionStatus::Absent;
    }
    if projects
        .iter()
        .any(|p| p.status == VbaRecognitionStatus::Partial)
    {
        return VbaRecognitionStatus::Partial;
    }
    if projects
        .iter()
        .any(|p| p.status == VbaRecognitionStatus::Error)
    {
        return VbaRecognitionStatus::Error;
    }
    if projects
        .iter()
        .all(|p| p.status == VbaRecognitionStatus::Extracted)
    {
        return VbaRecognitionStatus::Extracted;
    }
    VbaRecognitionStatus::Recognized
}

fn aggregate_module_status(modules: &[VbaModuleReport]) -> VbaRecognitionStatus {
    if modules.is_empty() {
        return VbaRecognitionStatus::Recognized;
    }
    if modules
        .iter()
        .any(|m| m.status == VbaRecognitionStatus::Error)
    {
        return VbaRecognitionStatus::Error;
    }
    if modules
        .iter()
        .any(|m| m.status == VbaRecognitionStatus::Partial)
    {
        return VbaRecognitionStatus::Partial;
    }
    if modules
        .iter()
        .all(|m| m.status == VbaRecognitionStatus::Extracted)
    {
        return VbaRecognitionStatus::Extracted;
    }
    VbaRecognitionStatus::Recognized
}

fn map_module_state(
    state: MacroExtractionState,
    errors: &[String],
    include_source: bool,
) -> VbaRecognitionStatus {
    match state {
        MacroExtractionState::Extracted if include_source => VbaRecognitionStatus::Extracted,
        MacroExtractionState::Extracted => VbaRecognitionStatus::Recognized,
        MacroExtractionState::Partial => VbaRecognitionStatus::Partial,
        MacroExtractionState::DecodeFailed => VbaRecognitionStatus::Error,
        MacroExtractionState::MissingStream if errors.is_empty() => {
            VbaRecognitionStatus::Recognized
        }
        MacroExtractionState::MissingStream => VbaRecognitionStatus::Partial,
        MacroExtractionState::NotRequested => VbaRecognitionStatus::Recognized,
    }
}

fn module_kind_name(kind: MacroModuleType) -> &'static str {
    match kind {
        MacroModuleType::Standard => "standard",
        MacroModuleType::Class => "class",
        MacroModuleType::UserForm => "form",
        MacroModuleType::Document => "document",
    }
}

fn format_suspicious_call(call: &SuspiciousCall) -> String {
    match call.line {
        Some(line) => format!("{}@{}", call.name, line),
        None => call.name.clone(),
    }
}

#[cfg(test)]
mod tests;
