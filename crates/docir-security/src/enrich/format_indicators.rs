use crate::make_indicator;
use docir_core::ir::IRNode;
use docir_core::security::{SecurityInfo, ThreatIndicator, ThreatIndicatorType, ThreatLevel};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;

use super::helpers::{
    activex_indicator_details, is_activex_ole, macro_project_details, ole_indicator_details,
    ole_location, push_ole_object_indicators, push_remote_external_ref_indicators,
};
use super::xlm::build_xlm_indicators;

pub(super) fn build_ooxml_indicators(
    store: &IrStore,
    security: &SecurityInfo,
) -> Vec<ThreatIndicator> {
    let mut indicators = Vec::new();

    if let Some(macro_id) = security.macro_project {
        let (indicator_type, severity, description, location) = macro_indicator(store, macro_id);
        indicators.push(make_indicator(
            indicator_type,
            severity,
            description,
            location,
            Some(macro_id),
        ));
    }

    let mut reported_activex_ids = std::collections::HashSet::new();
    indicators.extend(ole_object_indicators(
        store,
        &security.ole_objects,
        &mut reported_activex_ids,
    ));
    indicators.extend(activex_control_indicators(
        store,
        &security.activex_controls,
        &reported_activex_ids,
    ));
    indicators.extend(build_xlm_indicators(store, security));

    indicators
}

pub(super) fn build_odf_indicators(
    store: &IrStore,
    security: &SecurityInfo,
) -> Vec<ThreatIndicator> {
    let mut indicators = Vec::new();

    push_remote_external_ref_indicators(store, &security.external_refs, &mut indicators);

    for _ in &security.ole_objects {
        indicators.push(make_indicator(
            ThreatIndicatorType::OleObject,
            ThreatLevel::High,
            "Embedded OLE object".to_string(),
            None,
            None,
        ));
    }

    for dde in &security.dde_fields {
        indicators.push(make_indicator(
            ThreatIndicatorType::DdeCommand,
            ThreatLevel::High,
            format!("DDE formula: {}", dde.instruction),
            dde.location.as_ref().map(|span| span.file_path.clone()),
            None,
        ));
    }

    indicators
}

pub(super) fn build_hwp_indicators(
    store: &IrStore,
    security: &SecurityInfo,
    hwpx_autoexec: bool,
) -> Vec<ThreatIndicator> {
    let mut indicators = Vec::new();

    push_remote_external_ref_indicators(store, &security.external_refs, &mut indicators);
    push_ole_object_indicators(
        store,
        &security.ole_objects,
        "Embedded OLE object",
        true,
        ole_location,
        &mut indicators,
    );

    if hwpx_autoexec
        && let Some(macro_id) = security.macro_project
        && matches!(
            store.get(macro_id),
            Some(IRNode::MacroProject(project)) if project.has_auto_exec
        )
    {
        indicators.push(make_indicator(
            ThreatIndicatorType::AutoExecMacro,
            ThreatLevel::Critical,
            "Auto-exec script detected".to_string(),
            None,
            Some(macro_id),
        ));
    }

    indicators
}

pub(super) fn build_rtf_indicators(
    store: &IrStore,
    security: &SecurityInfo,
) -> Vec<ThreatIndicator> {
    let mut indicators = Vec::new();

    push_remote_external_ref_indicators(store, &security.external_refs, &mut indicators);
    push_ole_object_indicators(
        store,
        &security.ole_objects,
        "Embedded OLE object",
        true,
        |_| Some("rtf".to_string()),
        &mut indicators,
    );

    indicators
}

fn macro_indicator(
    store: &IrStore,
    macro_id: NodeId,
) -> (ThreatIndicatorType, ThreatLevel, String, Option<String>) {
    match store.get(macro_id) {
        Some(IRNode::MacroProject(project)) => {
            let details = macro_project_details(project);
            if project.has_auto_exec {
                (
                    ThreatIndicatorType::AutoExecMacro,
                    ThreatLevel::Critical,
                    details.1,
                    details.0,
                )
            } else {
                (
                    ThreatIndicatorType::MacroProject,
                    ThreatLevel::High,
                    details.1,
                    details.0,
                )
            }
        }
        _ => (
            ThreatIndicatorType::MacroProject,
            ThreatLevel::High,
            "VBA macro project found".to_string(),
            None,
        ),
    }
}

fn ole_object_indicators(
    store: &IrStore,
    ole_objects: &[NodeId],
    reported_activex_ids: &mut std::collections::HashSet<NodeId>,
) -> Vec<ThreatIndicator> {
    let mut indicators = Vec::new();
    for id in ole_objects {
        let Some(IRNode::OleObject(ole)) = store.get(*id) else {
            continue;
        };
        if is_activex_ole(ole) {
            reported_activex_ids.insert(*id);
            let (location, description) = ole_indicator_details(
                "ActiveX control binary found at",
                "ActiveX control binary found",
                ole,
            );
            indicators.push(make_indicator(
                ThreatIndicatorType::ActiveXControl,
                ThreatLevel::High,
                description,
                location,
                Some(*id),
            ));
        } else {
            let (location, description) =
                ole_indicator_details("OLE object found at", "OLE object found", ole);
            indicators.push(make_indicator(
                ThreatIndicatorType::OleObject,
                ThreatLevel::High,
                description,
                location,
                Some(*id),
            ));
        }
    }
    indicators
}

fn activex_control_indicators(
    store: &IrStore,
    activex_controls: &[NodeId],
    reported_activex_ids: &std::collections::HashSet<NodeId>,
) -> Vec<ThreatIndicator> {
    let mut indicators = Vec::new();
    for id in activex_controls {
        if reported_activex_ids.contains(id) {
            continue;
        }
        let Some(IRNode::ActiveXControl(control)) = store.get(*id) else {
            continue;
        };
        let (location, description) = activex_indicator_details(control);
        indicators.push(make_indicator(
            ThreatIndicatorType::ActiveXControl,
            ThreatLevel::High,
            description,
            location,
            Some(*id),
        ));
    }
    indicators
}
