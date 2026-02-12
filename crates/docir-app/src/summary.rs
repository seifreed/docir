use crate::ParsedDocument;
use docir_core::ir::IRNode;
use docir_core::visitor::{IrVisitor, NodeCounter, PreOrderWalker, VisitControl, VisitorResult};

#[derive(Debug, Clone)]
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
    pub threat_level: String,
    pub has_macro_project: bool,
    pub ole_objects: usize,
    pub external_refs: usize,
    pub dde_fields: usize,
    pub activex_controls: usize,
    pub xlm_macros: usize,
}

#[derive(Debug, Clone)]
pub struct ThreatIndicatorSummary {
    pub severity: String,
    pub indicator_type: String,
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

pub fn summarize_document(parsed: &ParsedDocument) -> Option<DocumentSummary> {
    let doc = parsed.document()?;

    let metadata = if let Some(meta_id) = doc.metadata {
        match parsed.store().get(meta_id) {
            Some(IRNode::Metadata(meta)) => MetadataSummary {
                title: meta.title.clone(),
                author: meta.creator.clone(),
                modified: meta.modified.clone(),
                application: meta.application.clone(),
            },
            _ => MetadataSummary {
                title: None,
                author: None,
                modified: None,
                application: None,
            },
        }
    } else {
        MetadataSummary {
            title: None,
            author: None,
            modified: None,
            application: None,
        }
    };

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

    let mut text_collector = TextStats::new();
    let mut walker = PreOrderWalker::new(parsed.store(), parsed.root_id());
    let _ = walker.walk(&mut text_collector);
    let text_stats = TextStatsSummary {
        char_count: text_collector.char_count,
        word_count: text_collector.word_count,
    };

    let metrics = parsed.metrics().map(|m| ParseMetricsSummary {
        content_types_ms: m.content_types_ms,
        relationships_ms: m.relationships_ms,
        main_parse_ms: m.main_parse_ms,
        shared_parts_ms: m.shared_parts_ms,
        security_scan_ms: m.security_scan_ms,
        extension_parts_ms: m.extension_parts_ms,
        normalization_ms: m.normalization_ms,
    });

    let security = SecuritySummary {
        threat_level: doc.security.threat_level.to_string(),
        has_macro_project: doc.security.macro_project.is_some(),
        ole_objects: doc.security.ole_objects.len(),
        external_refs: doc.security.external_refs.len(),
        dde_fields: doc.security.dde_fields.len(),
        activex_controls: doc.security.activex_controls.len(),
        xlm_macros: doc.security.xlm_macros.len(),
    };

    let threat_indicators = doc
        .security
        .threat_indicators
        .iter()
        .map(|indicator| ThreatIndicatorSummary {
            severity: indicator.severity.to_string(),
            indicator_type: format!("{:?}", indicator.indicator_type),
            description: indicator.description.clone(),
        })
        .collect();

    Some(DocumentSummary {
        format: doc.format.display_name().to_string(),
        metadata,
        node_counts: counts,
        text_stats,
        metrics,
        security,
        threat_indicators,
    })
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
