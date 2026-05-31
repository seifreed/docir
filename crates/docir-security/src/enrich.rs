use docir_core::ir::{Field, IRNode};
use docir_core::security::DdeField;
use docir_core::types::{DocumentFormat, NodeId};
use docir_core::visitor::IrStore;

mod dde;
mod format_indicators;
mod helpers;
mod xlm;

use self::dde::parse_dde_instruction;
use self::format_indicators::{
    build_hwp_indicators, build_odf_indicators, build_ooxml_indicators, build_rtf_indicators,
};
use self::xlm::apply_xlm_defined_name_targets;

/// Public API entrypoint: populate_security_indicators.
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
        DocumentFormat::Hwp => build_hwp_indicators(store, &security, true),
        DocumentFormat::Hwpx => build_hwp_indicators(store, &security, true),
        DocumentFormat::Rtf => build_rtf_indicators(store, &security),
    };
    indicators.append(&mut generated);

    if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
        doc.security.apply_scan_result(security, indicators);
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

    let existing_ole: std::collections::HashSet<NodeId> =
        security.ole_objects.iter().copied().collect();
    let store_ole: Vec<NodeId> = store
        .iter()
        .filter_map(|(id, node)| match node {
            IRNode::OleObject(_) => Some(*id),
            _ => None,
        })
        .collect();
    let missing_ole: Vec<NodeId> = store_ole
        .into_iter()
        .filter(|id| !existing_ole.contains(id))
        .collect();
    security.ole_objects.extend(missing_ole);

    let existing_refs: std::collections::HashSet<NodeId> =
        security.external_refs.iter().copied().collect();
    let store_refs: Vec<NodeId> = store
        .iter()
        .filter_map(|(id, node)| match node {
            IRNode::ExternalReference(_) => Some(*id),
            _ => None,
        })
        .collect();
    let missing_refs: Vec<NodeId> = store_refs
        .into_iter()
        .filter(|id| !existing_refs.contains(id))
        .collect();
    security.external_refs.extend(missing_refs);

    let existing_activex: std::collections::HashSet<NodeId> =
        security.activex_controls.iter().copied().collect();
    let store_activex: Vec<NodeId> = store
        .iter()
        .filter_map(|(id, node)| match node {
            IRNode::ActiveXControl(_) => Some(*id),
            _ => None,
        })
        .collect();
    let missing_activex: Vec<NodeId> = store_activex
        .into_iter()
        .filter(|id| !existing_activex.contains(id))
        .collect();
    security.activex_controls.extend(missing_activex);

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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::make_indicator;
    use docir_core::ir::{Document, Field, IRNode};
    use docir_core::security::{
        ActiveXControl, ExternalRefType, ExternalReference, MacroProject, OleObject,
        ThreatIndicatorType, ThreatLevel,
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
        assert!(
            indicators
                .iter()
                .any(|i| i.indicator_type == ThreatIndicatorType::RemoteResource)
        );
        assert!(indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::DdeCommand
                && i.description.contains("DDE formula")
                && i.location.as_deref() == Some("content.xml")
        }));
        // file:// is now also flagged as remote (security improvement)
        assert!(indicators.iter().any(|i| {
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

    #[test]
    fn populate_security_indicators_ooxml_adds_macro_and_activex_control_indicators() {
        let mut store = IrStore::new();
        let doc = Document::new(DocumentFormat::Spreadsheet);
        let root_id = doc.id;

        let mut project = MacroProject::new();
        project.has_auto_exec = true;

        let mut control = ActiveXControl::new();
        control.name = Some("Button1".to_string());

        store.insert(IRNode::Document(doc));
        store.insert(IRNode::MacroProject(project));
        store.insert(IRNode::ActiveXControl(control));

        populate_security_indicators(&mut store, root_id);

        let Some(IRNode::Document(doc)) = store.get(root_id) else {
            panic!("missing document");
        };
        assert!(doc.security.threat_indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::AutoExecMacro
                && i.description.contains("VBA macro project found")
        }));
        assert!(doc.security.threat_indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::ActiveXControl
                && i.description.contains("ActiveX control found")
        }));
    }

    #[test]
    fn populate_security_indicators_hwpx_emits_autoexec_but_hwp_does_not() {
        fn run_for_format(format: DocumentFormat) -> docir_core::security::SecurityInfo {
            let mut store = IrStore::new();
            let doc = Document::new(format);
            let root_id = doc.id;

            let mut project = MacroProject::new();
            project.has_auto_exec = true;

            let remote =
                ExternalReference::new(ExternalRefType::DataConnection, "https://evil.test/data");
            let mut ole = OleObject::new();
            ole.name = Some("EmbeddedObject".to_string());

            store.insert(IRNode::Document(doc));
            store.insert(IRNode::MacroProject(project));
            store.insert(IRNode::ExternalReference(remote));
            store.insert(IRNode::OleObject(ole));

            populate_security_indicators(&mut store, root_id);

            let Some(IRNode::Document(doc)) = store.get(root_id) else {
                panic!("missing document");
            };
            doc.security.clone()
        }

        let hwpx = run_for_format(DocumentFormat::Hwpx);
        assert!(
            hwpx.threat_indicators
                .iter()
                .any(|i| i.indicator_type == ThreatIndicatorType::AutoExecMacro)
        );
        assert!(hwpx.threat_indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::RemoteResource
                && i.description.contains("https://evil.test/data")
        }));
        assert!(hwpx.threat_indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::OleObject
                && i.location.as_deref() == Some("EmbeddedObject")
        }));

        let hwp = run_for_format(DocumentFormat::Hwp);
        assert!(
            hwp.threat_indicators
                .iter()
                .any(|i| i.indicator_type == ThreatIndicatorType::AutoExecMacro)
        );
        assert!(
            hwp.threat_indicators
                .iter()
                .any(|i| i.indicator_type == ThreatIndicatorType::OleObject)
        );
    }

    #[test]
    fn populate_security_indicators_rtf_preserves_existing_indicators_and_locations() {
        let mut store = IrStore::new();
        let mut doc = Document::new(DocumentFormat::Rtf);
        let root_id = doc.id;
        doc.security.threat_indicators.push(make_indicator(
            ThreatIndicatorType::SuspiciousLink,
            ThreatLevel::Low,
            "existing indicator".to_string(),
            None,
            None,
        ));

        let ole = OleObject::new();
        let remote = ExternalReference::new(ExternalRefType::Image, "https://evil.test/a.png");

        store.insert(IRNode::Document(doc));
        store.insert(IRNode::OleObject(ole));
        store.insert(IRNode::ExternalReference(remote));

        populate_security_indicators(&mut store, root_id);

        let Some(IRNode::Document(doc)) = store.get(root_id) else {
            panic!("missing document");
        };
        assert!(
            doc.security
                .threat_indicators
                .iter()
                .any(|i| i.description == "existing indicator")
        );
        assert!(doc.security.threat_indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::OleObject
                && i.location.as_deref() == Some("rtf")
        }));
        assert!(doc.security.threat_indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::RemoteResource
                && i.description.contains("https://evil.test/a.png")
        }));
    }
}
