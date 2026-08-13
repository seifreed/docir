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
    security.ole_objects.sort_unstable_by_key(NodeId::as_u64);

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
    security.external_refs.sort_unstable_by_key(NodeId::as_u64);

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
    security
        .activex_controls
        .sort_unstable_by_key(NodeId::as_u64);

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
    out.sort_unstable_by(|left, right| {
        let left_location = left.location.as_ref();
        let right_location = right.location.as_ref();
        left_location
            .map(|span| (&span.file_path, span.line, span.column))
            .cmp(&right_location.map(|span| (&span.file_path, span.line, span.column)))
            .then_with(|| left.instruction.cmp(&right.instruction))
    });
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
    fn populate_security_indicators_orders_rebuilt_security_collections() {
        let mut store = IrStore::new();
        let doc = Document::new(DocumentFormat::WordProcessing);
        let root_id = doc.id;
        let mut ole_first = OleObject::new();
        ole_first.span = Some(SourceSpan::new("word/embeddings/first.bin"));
        let mut ole_second = OleObject::new();
        ole_second.span = Some(SourceSpan::new("word/embeddings/second.bin"));
        let ole_ids = [ole_first.id, ole_second.id];
        let mut external_first =
            ExternalReference::new(ExternalRefType::Hyperlink, "https://first.test");
        external_first.span = Some(SourceSpan::new("word/document.xml"));
        let mut external_second =
            ExternalReference::new(ExternalRefType::Hyperlink, "https://second.test");
        external_second.span = Some(SourceSpan::new("word/header1.xml"));
        let external_ids = [external_first.id, external_second.id];
        let mut activex_first = ActiveXControl::new();
        activex_first.span = Some(SourceSpan::new("word/activeX/first.bin"));
        let mut activex_second = ActiveXControl::new();
        activex_second.span = Some(SourceSpan::new("word/activeX/second.bin"));
        let activex_ids = [activex_first.id, activex_second.id];
        let mut dde_first = Field::new(Some(r#"DDEAUTO "cmd" "/c first" "A1""#.to_string()));
        dde_first.span = Some(SourceSpan::new("z.xml"));
        let mut dde_second = Field::new(Some(r#"DDEAUTO "cmd" "/c second" "A1""#.to_string()));
        dde_second.span = Some(SourceSpan::new("a.xml"));

        store.insert(IRNode::Document(doc));
        store.insert(IRNode::OleObject(ole_second));
        store.insert(IRNode::OleObject(ole_first));
        store.insert(IRNode::ExternalReference(external_second));
        store.insert(IRNode::ExternalReference(external_first));
        store.insert(IRNode::ActiveXControl(activex_second));
        store.insert(IRNode::ActiveXControl(activex_first));
        store.insert(IRNode::Field(dde_first));
        store.insert(IRNode::Field(dde_second));

        populate_security_indicators(&mut store, root_id);

        let Some(IRNode::Document(doc)) = store.get(root_id) else {
            panic!("missing document");
        };
        let mut expected_ole = ole_ids.to_vec();
        expected_ole.sort_unstable_by_key(NodeId::as_u64);
        let mut expected_external = external_ids.to_vec();
        expected_external.sort_unstable_by_key(NodeId::as_u64);
        let mut expected_activex = activex_ids.to_vec();
        expected_activex.sort_unstable_by_key(NodeId::as_u64);
        assert_eq!(doc.security.ole_objects, expected_ole);
        assert_eq!(doc.security.external_refs, expected_external);
        assert_eq!(doc.security.activex_controls, expected_activex);
        assert_eq!(
            doc.security.dde_fields[0]
                .location
                .as_ref()
                .unwrap()
                .file_path,
            "a.xml"
        );
        assert_eq!(
            doc.security.dde_fields[1]
                .location
                .as_ref()
                .unwrap()
                .file_path,
            "z.xml"
        );
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
