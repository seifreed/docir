use crate::make_indicator;
use docir_core::ir::{Field, IRNode};
use docir_core::security::{DdeField, ThreatIndicator, ThreatIndicatorType, ThreatLevel};
use docir_core::types::{DocumentFormat, NodeId};
use docir_core::visitor::IrStore;

mod dde;
mod helpers;
mod xlm;

use self::dde::parse_dde_instruction;
use self::helpers::{
    activex_indicator_details, is_activex_ole, macro_project_details, ole_indicator_details,
    ole_location, push_ole_object_indicators, push_remote_external_ref_indicators,
};
use self::xlm::{apply_xlm_defined_name_targets, build_xlm_indicators};

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
            if matches!(
                store.get(macro_id),
                Some(IRNode::MacroProject(project)) if project.has_auto_exec
            ) {
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

    indicators
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

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::{Document, Field, IRNode};
    use docir_core::security::{
        ExternalRefType, ExternalReference, OleObject, ThreatIndicatorType,
    };
    use docir_core::types::{DocumentFormat, SourceSpan};

    #[test]
    fn populate_security_indicators_generates_dde_and_remote_ref_for_odf() {
        let mut store = IrStore::new();
        let doc = Document::new(DocumentFormat::OdfText);
        let root_id = doc.id;

        let mut remote = ExternalReference::new(ExternalRefType::Hyperlink, "https://evil.test");
        remote.span = Some(SourceSpan::new("content.xml"));
        let local = ExternalReference::new(ExternalRefType::Hyperlink, "file:///tmp/local");

        let mut dde_field = Field::new(Some(r#"DDEAUTO "cmd" "/c calc" "A1""#.to_string()));
        dde_field.span = Some(SourceSpan::new("content.xml"));

        store.insert(IRNode::Document(doc));
        store.insert(IRNode::ExternalReference(remote));
        store.insert(IRNode::ExternalReference(local));
        store.insert(IRNode::Field(dde_field));

        populate_security_indicators(&mut store, root_id);

        let Some(IRNode::Document(doc)) = store.get(root_id) else {
            panic!("missing document");
        };
        let indicators = &doc.security.threat_indicators;
        assert!(indicators
            .iter()
            .any(|i| i.indicator_type == ThreatIndicatorType::RemoteResource));
        assert!(indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::DdeCommand
                && i.description.contains("DDE formula")
                && i.location.as_deref() == Some("content.xml")
        }));
        assert!(!indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::RemoteResource
                && i.description.contains("file:///tmp/local")
        }));
    }

    #[test]
    fn populate_security_indicators_shapes_activex_and_ole_locations() {
        let mut store = IrStore::new();
        let doc = Document::new(DocumentFormat::WordProcessing);
        let root_id = doc.id;

        let mut activex_ole = OleObject::new();
        activex_ole.span = Some(SourceSpan::new("word/activeX/activeX1.bin"));
        let mut regular_ole = OleObject::new();
        regular_ole.span = Some(SourceSpan::new("word/embeddings/object1.bin"));

        store.insert(IRNode::Document(doc));
        store.insert(IRNode::OleObject(activex_ole));
        store.insert(IRNode::OleObject(regular_ole));

        populate_security_indicators(&mut store, root_id);

        let Some(IRNode::Document(doc)) = store.get(root_id) else {
            panic!("missing document");
        };
        assert!(doc.security.threat_indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::ActiveXControl
                && i.description.contains("ActiveX control binary found at")
                && i.location.as_deref() == Some("word/activeX/activeX1.bin")
        }));
        assert!(doc.security.threat_indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::OleObject
                && i.description.contains("OLE object found at")
                && i.location.as_deref() == Some("word/embeddings/object1.bin")
        }));
    }
}
