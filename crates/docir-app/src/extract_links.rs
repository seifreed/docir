use crate::ParsedDocument;
use docir_core::security::{DdeFieldType, ThreatLevel};
use serde::Serialize;

/// Dedicated DDE/link extraction report for analyst workflows.
#[derive(Debug, Clone, Serialize)]
pub struct LinkExtractionReport {
    pub document_format: String,
    pub link_count: usize,
    pub links: Vec<LinkArtifact>,
}

/// A single extracted link-like active content instruction.
#[derive(Debug, Clone, Serialize)]
pub struct LinkArtifact {
    pub kind: String,
    pub risk: ThreatLevel,
    pub raw_text: String,
    pub normalized: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub location: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub application: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub topic: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub item: Option<String>,
}

impl LinkExtractionReport {
    /// Builds a dedicated link extraction report from the parsed document.
    pub fn from_parsed(parsed: &ParsedDocument) -> Self {
        let mut links = Vec::new();
        if let Some(security) = parsed.security_info() {
            for dde in &security.dde_fields {
                links.push(LinkArtifact {
                    kind: match dde.field_type {
                        DdeFieldType::Dde => "dde".to_string(),
                        DdeFieldType::DdeAuto => "ddeauto".to_string(),
                    },
                    risk: ThreatLevel::High,
                    raw_text: dde.instruction.clone(),
                    normalized: normalize_dde(dde),
                    location: dde.location.as_ref().map(|span| span.file_path.clone()),
                    application: non_empty(&dde.application),
                    topic: dde.topic.clone(),
                    item: dde.item.clone(),
                });
            }
        }

        Self {
            document_format: parsed.format().extension().to_string(),
            link_count: links.len(),
            links,
        }
    }
}

fn normalize_dde(dde: &docir_core::security::DdeField) -> String {
    let kind = match dde.field_type {
        DdeFieldType::Dde => "DDE",
        DdeFieldType::DdeAuto => "DDEAUTO",
    };
    let application = non_empty(&dde.application).unwrap_or_else(|| "<unknown>".to_string());
    let topic = dde.topic.clone().unwrap_or_else(|| "<none>".to_string());
    let item = dde.item.clone().unwrap_or_else(|| "<none>".to_string());
    format!("{kind} app={application} topic={topic} item={item}")
}

fn non_empty(value: &str) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::LinkExtractionReport;
    use crate::ParsedDocument;
    use docir_core::ir::{Document, IRNode};
    use docir_core::security::{DdeField, DdeFieldType, ThreatLevel};
    use docir_core::types::{DocumentFormat, SourceSpan};
    use docir_core::visitor::IrStore;

    #[test]
    fn extract_links_collects_dde_fields() {
        let mut store = IrStore::new();
        let mut doc = Document::new(DocumentFormat::WordProcessing);
        doc.security.dde_fields.push(DdeField {
            field_type: DdeFieldType::DdeAuto,
            application: "cmd".to_string(),
            topic: Some("/c calc".to_string()),
            item: Some("A1".to_string()),
            instruction: r#"DDEAUTO "cmd" "/c calc" "A1""#.to_string(),
            location: Some(SourceSpan::new("word/document.xml")),
        });
        let root_id = doc.id;
        store.insert(IRNode::Document(doc));
        let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
            root_id,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        });

        let report = LinkExtractionReport::from_parsed(&parsed);
        assert_eq!(report.link_count, 1);
        assert_eq!(report.links[0].kind, "ddeauto");
        assert_eq!(report.links[0].risk, ThreatLevel::High);
        assert_eq!(
            report.links[0].normalized,
            "DDEAUTO app=cmd topic=/c calc item=A1"
        );
        assert_eq!(
            report.links[0].location.as_deref(),
            Some("word/document.xml")
        );
    }

    #[test]
    fn extract_links_reports_empty_when_no_dde_exists() {
        let mut store = IrStore::new();
        let doc = Document::new(DocumentFormat::WordProcessing);
        let root_id = doc.id;
        store.insert(IRNode::Document(doc));
        let parsed = ParsedDocument::new(docir_parser::parser::ParsedDocument {
            root_id,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        });

        let report = LinkExtractionReport::from_parsed(&parsed);
        assert_eq!(report.link_count, 0);
        assert!(report.links.is_empty());
    }
}
