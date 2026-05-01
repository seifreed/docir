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
        .or(ole.name.clone())
}

pub(super) fn is_activex_ole(ole: &OleObject) -> bool {
    let path = ole
        .span
        .as_ref()
        .map(|s| s.file_path.as_str())
        .or(ole.name.as_deref())
        .unwrap_or("");
    let path_lower = path.to_ascii_lowercase();
    ole.class_id.is_some() || path_lower.contains("activex") || path_lower.ends_with(".ocx")
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::IRNode;
    use docir_core::security::{
        ExternalRefType, ExternalReference, OleObject, ThreatIndicatorType,
    };
    use docir_core::types::{NodeId, SourceSpan};
    use docir_core::visitor::IrStore;

    #[test]
    fn push_remote_external_ref_indicators_includes_file_protocol() {
        let mut store = IrStore::new();
        let mut remote = ExternalReference::new(ExternalRefType::Hyperlink, "https://evil.test");
        remote.span = Some(SourceSpan::new("word/_rels/document.xml.rels"));
        let local_file = ExternalReference::new(ExternalRefType::Hyperlink, "file:///tmp/report");
        let missing = NodeId::new();
        let refs = vec![remote.id, local_file.id, missing];
        store.insert(IRNode::ExternalReference(remote));
        store.insert(IRNode::ExternalReference(local_file));

        let mut indicators = Vec::new();
        push_remote_external_ref_indicators(&store, &refs, &mut indicators);
        // Both https:// and file:// are now flagged as remote
        assert_eq!(indicators.len(), 2);
        assert_eq!(
            indicators[0].indicator_type,
            ThreatIndicatorType::RemoteResource
        );
        assert!(indicators[0].description.contains("https://evil.test"));
        assert!(indicators[1].description.contains("file:///tmp/report"));
    }

    #[test]
    fn ole_indicator_details_prefers_span_then_name_then_fallback() {
        let mut with_span = OleObject::new();
        with_span.span = Some(SourceSpan::new("word/embeddings/ole1.bin"));
        let (loc, desc) =
            ole_indicator_details("OLE object found at", "OLE object found", &with_span);
        assert_eq!(loc.as_deref(), Some("word/embeddings/ole1.bin"));
        assert!(desc.contains("word/embeddings/ole1.bin"));

        let mut with_name = OleObject::new();
        with_name.name = Some("ObjectName".to_string());
        let (loc, desc) =
            ole_indicator_details("OLE object found at", "OLE object found", &with_name);
        assert_eq!(loc.as_deref(), Some("ObjectName"));
        assert!(desc.contains("ObjectName"));

        let without_loc = OleObject::new();
        let (loc, desc) =
            ole_indicator_details("OLE object found at", "OLE object found", &without_loc);
        assert!(loc.is_none());
        assert_eq!(desc, "OLE object found");
    }

    #[test]
    fn is_activex_ole_matches_span_or_name_case_insensitively() {
        let mut by_span = OleObject::new();
        by_span.span = Some(SourceSpan::new("word/activeX/activeX1.bin"));
        assert!(is_activex_ole(&by_span));

        let mut by_name = OleObject::new();
        by_name.name = Some("PPT/ACTIVEX/BINARY.BIN".to_string());
        assert!(is_activex_ole(&by_name));

        let other = OleObject::new();
        assert!(!is_activex_ole(&other));
    }
}
