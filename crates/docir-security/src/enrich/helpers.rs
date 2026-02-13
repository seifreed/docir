use crate::make_indicator;
use docir_core::ir::IRNode;
use docir_core::security::{
    ActiveXControl, MacroProject, OleObject, ThreatIndicator, ThreatIndicatorType, ThreatLevel,
};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;

pub(super) fn push_remote_external_ref_indicators(
    store: &IrStore,
    external_refs: &[NodeId],
    indicators: &mut Vec<ThreatIndicator>,
) {
    for id in external_refs {
        let Some(IRNode::ExternalReference(ext)) = store.get(*id) else {
            continue;
        };
        if ext.is_remote() {
            indicators.push(make_indicator(
                ThreatIndicatorType::RemoteResource,
                ThreatLevel::Medium,
                format!("Remote resource: {}", ext.target),
                ext.span.as_ref().map(|s| s.file_path.clone()),
                None,
            ));
        }
    }
}

pub(super) fn push_ole_object_indicators<F>(
    store: &IrStore,
    ole_objects: &[NodeId],
    description: &str,
    include_node_id: bool,
    location_fn: F,
    indicators: &mut Vec<ThreatIndicator>,
) where
    F: Fn(&OleObject) -> Option<String>,
{
    for id in ole_objects {
        let Some(IRNode::OleObject(ole)) = store.get(*id) else {
            continue;
        };
        let node_id = if include_node_id { Some(*id) } else { None };
        indicators.push(make_indicator(
            ThreatIndicatorType::OleObject,
            ThreatLevel::High,
            description.to_string(),
            location_fn(ole),
            node_id,
        ));
    }
}

pub(super) fn macro_project_details(project: &MacroProject) -> (Option<String>, String) {
    if let Some(span) = project.span.as_ref() {
        (
            Some(span.file_path.clone()),
            format!("VBA macro project found at {}", span.file_path),
        )
    } else {
        (None, "VBA macro project found".to_string())
    }
}

pub(super) fn activex_indicator_details(control: &ActiveXControl) -> (Option<String>, String) {
    if let Some(span) = control.span.as_ref() {
        (
            Some(span.file_path.clone()),
            format!("ActiveX control found at {}", span.file_path),
        )
    } else {
        (None, "ActiveX control found".to_string())
    }
}

pub(super) fn ole_indicator_details(
    prefix: &str,
    fallback: &str,
    ole: &OleObject,
) -> (Option<String>, String) {
    if let Some(span) = ole.span.as_ref() {
        (
            Some(span.file_path.clone()),
            format!("{prefix} {}", span.file_path),
        )
    } else if let Some(name) = ole.name.as_ref() {
        (Some(name.clone()), format!("{prefix} {}", name))
    } else {
        (None, fallback.to_string())
    }
}

pub(super) fn ole_location(ole: &OleObject) -> Option<String> {
    ole.span
        .as_ref()
        .map(|s| s.file_path.clone())
        .or_else(|| ole.name.clone())
}

pub(super) fn is_activex_ole(ole: &OleObject) -> bool {
    let path = ole
        .span
        .as_ref()
        .map(|s| s.file_path.as_str())
        .or_else(|| ole.name.as_deref())
        .unwrap_or("");
    path.to_ascii_lowercase().contains("activex")
}
