use crate::VbaRecognitionReport;
use docir_core::security::ThreatLevel;

use super::helpers::boolean_or_count_indicator;
use super::DocumentIndicator;

pub(super) fn collect_macro_indicators(vba: &VbaRecognitionReport) -> Vec<DocumentIndicator> {
    let mut indicators = Vec::with_capacity(3);

    let macro_count = vba.projects.len();
    let macro_evidence = vba
        .projects
        .iter()
        .map(|project| {
            project
                .container_path
                .clone()
                .or_else(|| project.storage_root.clone())
                .or_else(|| project.project_name.clone())
                .unwrap_or_else(|| project.node_id.clone())
        })
        .collect::<Vec<_>>();
    indicators.push(boolean_or_count_indicator(
        "macros",
        macro_count,
        ThreatLevel::Critical,
        "VBA project or macro-capable content detected",
        "No macro-capable content detected",
        macro_evidence,
    ));

    let auto_exec = vba
        .projects
        .iter()
        .flat_map(|project| {
            project.auto_exec_procedures.iter().map(|proc_name| {
                format!(
                    "{}:{}",
                    project
                        .project_name
                        .clone()
                        .unwrap_or_else(|| project.node_id.clone()),
                    proc_name
                )
            })
        })
        .collect::<Vec<_>>();
    indicators.push(boolean_or_count_indicator(
        "autoexec",
        auto_exec.len(),
        ThreatLevel::Critical,
        "Auto-execute VBA entrypoints detected",
        "No auto-execute VBA entrypoints detected",
        auto_exec,
    ));

    let protected_projects = vba
        .projects
        .iter()
        .filter(|project| project.is_protected)
        .map(|project| {
            project
                .project_name
                .clone()
                .or_else(|| project.storage_root.clone())
                .unwrap_or_else(|| project.node_id.clone())
        })
        .collect::<Vec<_>>();
    indicators.push(boolean_or_count_indicator(
        "protected-vba",
        protected_projects.len(),
        ThreatLevel::Medium,
        "Password-protected VBA projects detected",
        "No password-protected VBA projects detected",
        protected_projects,
    ));

    indicators
}
