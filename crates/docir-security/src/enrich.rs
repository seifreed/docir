use crate::make_indicator;
use docir_core::ir::IRNode;
use docir_core::security::{
    ActiveXControl, MacroProject, OleObject, ThreatIndicator, ThreatIndicatorType, ThreatLevel,
};
use docir_core::types::{DocumentFormat, NodeId};
use docir_core::visitor::IrStore;

pub fn populate_security_indicators(store: &mut IrStore, root_id: NodeId) {
    let (format, security) = match store.get(root_id) {
        Some(IRNode::Document(doc)) => (doc.format, doc.security.clone()),
        _ => return,
    };

    let mut indicators = security.threat_indicators.clone();
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
        doc.security.threat_indicators = indicators;
        doc.security.recalculate_threat_level();
    }
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

    indicators
}

fn build_odf_indicators(
    store: &IrStore,
    security: &docir_core::security::SecurityInfo,
) -> Vec<ThreatIndicator> {
    let mut indicators = Vec::new();

    for id in &security.external_refs {
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

    for id in &security.external_refs {
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

    for id in &security.ole_objects {
        let Some(IRNode::OleObject(ole)) = store.get(*id) else {
            continue;
        };
        let location = ole_location(ole);
        indicators.push(make_indicator(
            ThreatIndicatorType::OleObject,
            ThreatLevel::High,
            "Embedded OLE object".to_string(),
            location,
            Some(*id),
        ));
    }

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

fn build_rtf_indicators(
    store: &IrStore,
    security: &docir_core::security::SecurityInfo,
) -> Vec<ThreatIndicator> {
    let mut indicators = Vec::new();

    for id in &security.external_refs {
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

    for id in &security.ole_objects {
        if store.get(*id).is_none() {
            continue;
        }
        indicators.push(make_indicator(
            ThreatIndicatorType::OleObject,
            ThreatLevel::High,
            "Embedded OLE object".to_string(),
            Some("rtf".to_string()),
            Some(*id),
        ));
    }

    indicators
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
