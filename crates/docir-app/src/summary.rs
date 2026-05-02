use crate::ParsedDocument;
use docir_core::ir::IRNode;
use docir_core::security::{ThreatIndicatorType, ThreatLevel};
use docir_core::visitor::{IrVisitor, NodeCounter, PreOrderWalker, VisitControl, VisitorResult};

#[derive(Debug, Clone, Default)]
pub struct MetadataSummary {
    pub title: Option<String>,
    pub author: Option<String>,
    pub modified: Option<String>,
    pub application: Option<String>,
}

#[derive(Debug, Clone)]
pub struct NodeCount {
    pub node_type: String,
    pub count: usize,
}

#[derive(Debug, Clone)]
pub struct TextStatsSummary {
    pub char_count: usize,
    pub word_count: usize,
}

#[derive(Debug, Clone)]
pub struct ParseMetricsSummary {
    pub content_types_ms: u128,
    pub relationships_ms: u128,
    pub main_parse_ms: u128,
    pub shared_parts_ms: u128,
    pub security_scan_ms: u128,
    pub extension_parts_ms: u128,
    pub normalization_ms: u128,
}

#[derive(Debug, Clone)]
pub struct SecuritySummary {
    pub threat_level: ThreatLevel,
    pub has_macro_project: bool,
    pub ole_objects: usize,
    pub external_refs: usize,
    pub dde_fields: usize,
    pub activex_controls: usize,
    pub xlm_macros: usize,
}

#[derive(Debug, Clone)]
pub struct ThreatIndicatorSummary {
    pub severity: ThreatLevel,
    pub indicator_type: ThreatIndicatorType,
    pub description: String,
}

#[derive(Debug, Clone)]
pub struct DocumentSummary {
    pub format: String,
    pub metadata: MetadataSummary,
    pub node_counts: Vec<NodeCount>,
    pub text_stats: TextStatsSummary,
    pub metrics: Option<ParseMetricsSummary>,
    pub security: SecuritySummary,
    pub threat_indicators: Vec<ThreatIndicatorSummary>,
}

/// Public API entrypoint: summarize_document.
pub fn summarize_document(parsed: &ParsedDocument) -> Option<DocumentSummary> {
    let doc = parsed.document()?;

    Some(DocumentSummary {
        format: doc.format.display_name().to_string(),
        metadata: build_metadata_summary(parsed, doc.metadata),
        node_counts: build_node_counts(parsed),
        text_stats: build_text_stats(parsed),
        metrics: build_metrics_summary(parsed),
        security: SecuritySummary::from(doc),
        threat_indicators: doc
            .security
            .threat_indicators
            .iter()
            .map(ThreatIndicatorSummary::from)
            .collect(),
    })
}

fn build_metadata_summary(
    parsed: &ParsedDocument,
    metadata_id: Option<docir_core::types::NodeId>,
) -> MetadataSummary {
    let Some(meta_id) = metadata_id else {
        return MetadataSummary::default();
    };
    match parsed.store().get(meta_id) {
        Some(IRNode::Metadata(meta)) => MetadataSummary {
            title: meta.title.clone(),
            author: meta.creator.clone(),
            modified: meta.modified.clone(),
            application: meta.application.clone(),
        },
        _ => MetadataSummary::default(),
    }
}

fn build_node_counts(parsed: &ParsedDocument) -> Vec<NodeCount> {
    let mut counter = NodeCounter::new();
    let mut walker = PreOrderWalker::new(parsed.store(), parsed.root_id());
    let _ = walker.walk(&mut counter);
    let mut counts: Vec<_> = counter
        .counts
        .iter()
        .map(|(node_type, count)| NodeCount {
            node_type: node_type.to_string(),
            count: *count,
        })
        .collect();
    counts.sort_by_key(|count| std::cmp::Reverse(count.count));
    counts
}

fn build_text_stats(parsed: &ParsedDocument) -> TextStatsSummary {
    let mut text_collector = TextStats::new();
    let mut walker = PreOrderWalker::new(parsed.store(), parsed.root_id());
    let _ = walker.walk(&mut text_collector);
    TextStatsSummary {
        char_count: text_collector.char_count,
        word_count: text_collector.word_count,
    }
}

fn build_metrics_summary(parsed: &ParsedDocument) -> Option<ParseMetricsSummary> {
    parsed.metrics().map(|m| ParseMetricsSummary {
        content_types_ms: m.content_types_ms,
        relationships_ms: m.relationships_ms,
        main_parse_ms: m.main_parse_ms,
        shared_parts_ms: m.shared_parts_ms,
        security_scan_ms: m.security_scan_ms,
        extension_parts_ms: m.extension_parts_ms,
        normalization_ms: m.normalization_ms,
    })
}

impl From<&docir_core::ir::Document> for SecuritySummary {
    fn from(doc: &docir_core::ir::Document) -> Self {
        SecuritySummary {
            threat_level: doc.security.threat_level,
            has_macro_project: doc.security.has_macro_project(),
            ole_objects: doc.security.ole_object_count(),
            external_refs: doc.security.external_ref_count(),
            dde_fields: doc.security.dde_field_count(),
            activex_controls: doc.security.activex_control_count(),
            xlm_macros: doc.security.xlm_macro_count(),
        }
    }
}

impl From<&docir_core::security::ThreatIndicator> for ThreatIndicatorSummary {
    fn from(indicator: &docir_core::security::ThreatIndicator) -> Self {
        ThreatIndicatorSummary {
            severity: indicator.severity,
            indicator_type: indicator.indicator_type,
            description: indicator.description.clone(),
        }
    }
}

struct TextStats {
    char_count: usize,
    word_count: usize,
}

impl TextStats {
    fn new() -> Self {
        Self {
            char_count: 0,
            word_count: 0,
        }
    }
}

impl IrVisitor for TextStats {
    fn visit_run(&mut self, run: &docir_core::ir::Run) -> VisitorResult<VisitControl> {
        self.char_count += run.text.chars().count();
        self.word_count += run.text.split_whitespace().count();
        Ok(VisitControl::Continue)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{ParseMetrics, ParsedDocument};
    use docir_core::ir::{Document, DocumentMetadata, IRNode, Paragraph, Run};
    use docir_core::security::{ThreatIndicator, ThreatIndicatorType, ThreatLevel};
    use docir_core::types::DocumentFormat;
    use docir_core::visitor::IrStore;

    fn parsed_with_document(include_metadata: bool) -> ParsedDocument {
        let mut store = IrStore::new();
        let mut doc = Document::new(DocumentFormat::WordProcessing);
        doc.security.threat_level = ThreatLevel::High;
        let external = docir_core::security::ExternalReference::new(
            docir_core::security::ExternalRefType::Hyperlink,
            "https://example.com",
        );
        let external_id = external.id;
        doc.security.external_refs.push(external_id);
        store.insert(IRNode::ExternalReference(external));
        doc.security.threat_indicators.push(ThreatIndicator {
            indicator_type: ThreatIndicatorType::ExternalTemplate,
            severity: ThreatLevel::High,
            description: "remote template".to_string(),
            location: Some("word/settings.xml".to_string()),
            node_id: None,
        });

        let run = Run::new("hello world");
        let run_id = run.id;
        let mut para = Paragraph::new();
        para.runs.push(run_id);
        let para_id = para.id;
        doc.content.push(para_id);

        if include_metadata {
            let mut meta = DocumentMetadata::new();
            meta.title = Some("Doc".to_string());
            meta.creator = Some("Author".to_string());
            meta.application = Some("app".to_string());
            let meta_id = meta.id;
            doc.metadata = Some(meta_id);
            store.insert(IRNode::Metadata(meta));
        }

        let root_id = doc.id;
        store.insert(IRNode::Run(run));
        store.insert(IRNode::Paragraph(para));
        store.insert(IRNode::Document(doc));

        ParsedDocument {
            inner: docir_parser::parser::ParsedDocument {
                root_id,
                format: DocumentFormat::WordProcessing,
                store,
                metrics: Some(ParseMetrics::default()),
            },
            metrics: Some(ParseMetrics::default()),
        }
    }

    #[test]
    fn summarize_document_returns_none_when_root_is_not_document() {
        let mut store = IrStore::new();
        let para = Paragraph::new();
        let root_id = para.id;
        store.insert(IRNode::Paragraph(para));
        let parsed = ParsedDocument {
            inner: docir_parser::parser::ParsedDocument {
                root_id,
                format: DocumentFormat::WordProcessing,
                store,
                metrics: None,
            },
            metrics: None,
        };
        assert!(summarize_document(&parsed).is_none());
    }

    #[test]
    fn summarize_document_builds_metadata_security_counts_and_text_stats() {
        let parsed = parsed_with_document(true);
        let summary = summarize_document(&parsed).expect("summary");

        assert_eq!(summary.format, "Word Document");
        assert_eq!(summary.metadata.title.as_deref(), Some("Doc"));
        assert_eq!(summary.metadata.author.as_deref(), Some("Author"));
        assert_eq!(summary.metadata.application.as_deref(), Some("app"));
        assert_eq!(summary.text_stats.word_count, 2);
        assert!(summary.text_stats.char_count >= 11);
        assert_eq!(summary.security.threat_level, ThreatLevel::High);
        assert_eq!(summary.security.external_refs, 1);
        assert_eq!(summary.threat_indicators.len(), 1);
        assert!(summary
            .node_counts
            .iter()
            .any(|count| count.node_type == "Document"));
        assert!(summary.metrics.is_some());
    }

    #[test]
    fn summarize_document_uses_default_metadata_when_missing_node() {
        let parsed = parsed_with_document(false);
        let summary = summarize_document(&parsed).expect("summary");
        assert!(summary.metadata.title.is_none());
        assert!(summary.metadata.author.is_none());
    }
}
