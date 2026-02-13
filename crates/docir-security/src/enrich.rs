use crate::make_indicator;
use docir_core::ir::{DefinedName, Field, IRNode};
use docir_core::security::{
    ActiveXControl, DdeField, MacroProject, OleObject, ThreatIndicator, ThreatIndicatorType,
    ThreatLevel,
};
use docir_core::types::{DocumentFormat, NodeId};
use docir_core::visitor::IrStore;
use std::collections::HashMap;

mod dde;

use self::dde::parse_dde_instruction;

pub fn populate_security_indicators(store: &mut IrStore, root_id: NodeId) {
    let (format, mut security) = match store.get(root_id) {
        Some(IRNode::Document(doc)) => (doc.format, doc.security.clone()),
        _ => return,
    };

    rebuild_security_info(store, &mut security);

    let mut indicators = security.threat_indicators.clone();
    apply_xlm_defined_name_targets(store, &mut security, &mut indicators);
    let mut generated = match format {
        DocumentFormat::WordProcessing
        | DocumentFormat::Spreadsheet
        | DocumentFormat::Presentation => build_ooxml_indicators(store, &security),
        DocumentFormat::OdfText
        | DocumentFormat::OdfSpreadsheet
        | DocumentFormat::OdfPresentation => build_odf_indicators(store, &security),
        DocumentFormat::Hwp => build_hwp_indicators(store, &security, false),
        DocumentFormat::Hwpx => build_hwp_indicators(store, &security, true),
        DocumentFormat::Rtf => build_rtf_indicators(store, &security),
    };
    indicators.append(&mut generated);

    if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
        doc.security.macro_project = security.macro_project;
        doc.security.ole_objects = security.ole_objects.clone();
        doc.security.external_refs = security.external_refs.clone();
        doc.security.activex_controls = security.activex_controls.clone();
        doc.security.dde_fields = security.dde_fields.clone();
        doc.security.threat_indicators = indicators;
        doc.security.recalculate_threat_level();
    }
}

fn rebuild_security_info(store: &IrStore, security: &mut docir_core::security::SecurityInfo) {
    if security.macro_project.is_none() {
        for (id, node) in store.iter() {
            if matches!(node, IRNode::MacroProject(_)) {
                security.macro_project = Some(*id);
                break;
            }
        }
    }

    if security.ole_objects.is_empty() {
        security.ole_objects = store
            .iter()
            .filter_map(|(id, node)| match node {
                IRNode::OleObject(_) => Some(*id),
                _ => None,
            })
            .collect();
    }

    if security.external_refs.is_empty() {
        security.external_refs = store
            .iter()
            .filter_map(|(id, node)| match node {
                IRNode::ExternalReference(_) => Some(*id),
                _ => None,
            })
            .collect();
    }

    if security.activex_controls.is_empty() {
        security.activex_controls = store
            .iter()
            .filter_map(|(id, node)| match node {
                IRNode::ActiveXControl(_) => Some(*id),
                _ => None,
            })
            .collect();
    }

    if security.dde_fields.is_empty() {
        security.dde_fields = scan_dde_fields(store);
    }
}

fn scan_dde_fields(store: &IrStore) -> Vec<DdeField> {
    let mut out = Vec::new();
    for node in store.values() {
        let IRNode::Field(Field {
            instruction: Some(instr),
            span,
            ..
        }) = node
        else {
            continue;
        };
        if let Some(mut dde) = parse_dde_instruction(instr) {
            dde.location = span.clone();
            out.push(dde);
        }
    }
    out
}

fn build_ooxml_indicators(
    store: &IrStore,
    security: &docir_core::security::SecurityInfo,
) -> Vec<ThreatIndicator> {
    let mut indicators = Vec::new();

    if let Some(macro_id) = security.macro_project {
        let (location, description) = match store.get(macro_id) {
            Some(IRNode::MacroProject(project)) => macro_project_details(project),
            _ => (None, "VBA macro project found".to_string()),
        };
        indicators.push(make_indicator(
            ThreatIndicatorType::AutoExecMacro,
            ThreatLevel::Critical,
            description,
            location,
            Some(macro_id),
        ));
    }

    for id in &security.ole_objects {
        let Some(IRNode::OleObject(ole)) = store.get(*id) else {
            continue;
        };
        if is_activex_ole(ole) {
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

    for id in &security.activex_controls {
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

    indicators.extend(build_xlm_indicators(store, security));

    indicators
}

fn build_odf_indicators(
    store: &IrStore,
    security: &docir_core::security::SecurityInfo,
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

fn build_hwp_indicators(
    store: &IrStore,
    security: &docir_core::security::SecurityInfo,
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

    if hwpx_autoexec {
        if let Some(macro_id) = security.macro_project {
            if let Some(IRNode::MacroProject(project)) = store.get(macro_id) {
                if project.has_auto_exec {
                    indicators.push(make_indicator(
                        ThreatIndicatorType::AutoExecMacro,
                        ThreatLevel::Critical,
                        "Auto-exec script detected".to_string(),
                        None,
                        Some(macro_id),
                    ));
                }
            }
        }
    }

    indicators
}

fn build_xlm_indicators(
    store: &IrStore,
    security: &docir_core::security::SecurityInfo,
) -> Vec<ThreatIndicator> {
    let mut indicators = Vec::new();
    let mut sheet_locations = HashMap::new();

    for node in store.values() {
        if let IRNode::Worksheet(sheet) = node {
            if let Some(span) = sheet.span.as_ref() {
                sheet_locations.insert(sheet.name.to_ascii_uppercase(), span.file_path.clone());
            }
        }
    }

    if security.xlm_macros.is_empty() {
        return indicators;
    }

    for macro_entry in &security.xlm_macros {
        let location = sheet_locations
            .get(&macro_entry.sheet_name.to_ascii_uppercase())
            .cloned();

        if macro_entry.has_auto_open {
            indicators.push(make_indicator(
                ThreatIndicatorType::XlmMacro,
                ThreatLevel::Critical,
                "XLM Auto_Open macro detected".to_string(),
                location.clone(),
                None,
            ));
        }

        if macro_entry.sheet_state != docir_core::ir::SheetState::Visible
            && !macro_entry.macro_cells.is_empty()
        {
            indicators.push(make_indicator(
                ThreatIndicatorType::HiddenMacroSheet,
                ThreatLevel::High,
                format!("Hidden macro sheet: {}", macro_entry.sheet_name),
                location.clone(),
                None,
            ));
        }

        for func in &macro_entry.dangerous_functions {
            indicators.push(make_indicator(
                ThreatIndicatorType::XlmMacro,
                ThreatLevel::Critical,
                format!("XLM macro function {} at {}", func.name, func.cell_ref),
                location.clone(),
                None,
            ));
        }
    }

    indicators
}

fn apply_xlm_defined_name_targets(
    store: &IrStore,
    security: &mut docir_core::security::SecurityInfo,
    indicators: &mut Vec<ThreatIndicator>,
) {
    let mut targets: Vec<Option<String>> = Vec::new();
    let mut location: Option<String> = None;

    for node in store.values() {
        if let IRNode::DefinedName(name) = node {
            if let Some(target) = auto_open_target_from_defined_name(name) {
                if location.is_none() {
                    location = name.span.as_ref().map(|s| s.file_path.clone());
                }
                targets.push(target);
            }
        }
    }

    if targets.is_empty() || security.xlm_macros.is_empty() {
        return;
    }

    let mut any_marked = false;
    for target in &targets {
        if let Some(target) = target {
            let target_upper = target.to_ascii_uppercase();
            for macro_entry in security.xlm_macros.iter_mut() {
                if macro_entry.sheet_name.to_ascii_uppercase() == target_upper {
                    macro_entry.has_auto_open = true;
                    any_marked = true;
                }
            }
        }
    }

    if !any_marked {
        for macro_entry in security.xlm_macros.iter_mut() {
            macro_entry.has_auto_open = true;
        }
    }

    indicators.push(make_indicator(
        ThreatIndicatorType::XlmMacro,
        ThreatLevel::Critical,
        "XLM Auto_Open defined name detected".to_string(),
        location,
        None,
    ));
}

fn auto_open_target_from_defined_name(name: &DefinedName) -> Option<Option<String>> {
    let upper = name.name.to_ascii_uppercase();
    if upper == "_XLNM.AUTO_OPEN" || upper == "AUTO_OPEN" || upper == "AUTO.OPEN" {
        let val = name.value.trim();
        if let Some((sheet, _)) = val.split_once('!') {
            let cleaned = sheet.trim().trim_matches('\'').to_string();
            if !cleaned.is_empty() {
                return Some(Some(cleaned));
            }
        }
        return Some(None);
    }
    None
}

fn build_rtf_indicators(
    store: &IrStore,
    security: &docir_core::security::SecurityInfo,
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

fn push_remote_external_ref_indicators(
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

fn push_ole_object_indicators<F>(
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

fn macro_project_details(project: &MacroProject) -> (Option<String>, String) {
    if let Some(span) = project.span.as_ref() {
        (
            Some(span.file_path.clone()),
            format!("VBA macro project found at {}", span.file_path),
        )
    } else {
        (None, "VBA macro project found".to_string())
    }
}

fn activex_indicator_details(control: &ActiveXControl) -> (Option<String>, String) {
    if let Some(span) = control.span.as_ref() {
        (
            Some(span.file_path.clone()),
            format!("ActiveX control found at {}", span.file_path),
        )
    } else {
        (None, "ActiveX control found".to_string())
    }
}

fn ole_indicator_details(
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

fn ole_location(ole: &OleObject) -> Option<String> {
    ole.span
        .as_ref()
        .map(|s| s.file_path.clone())
        .or_else(|| ole.name.clone())
}

fn is_activex_ole(ole: &OleObject) -> bool {
    let path = ole
        .span
        .as_ref()
        .map(|s| s.file_path.as_str())
        .or_else(|| ole.name.as_deref())
        .unwrap_or("");
    path.to_ascii_lowercase().contains("activex")
}
