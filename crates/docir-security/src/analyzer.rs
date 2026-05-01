//! Security analyzer for IR documents.

use crate::make_indicator;
use docir_core::ir::IRNode;
use docir_core::security::{SecurityInfo, ThreatIndicator, ThreatIndicatorType, ThreatLevel};
use docir_core::visitor::{IrStore, IrVisitor, PreOrderWalker, VisitControl, VisitorResult};

/// Security analyzer that examines an IR document for threats.
pub struct SecurityAnalyzer {
    findings: Vec<ThreatIndicator>,
}

impl SecurityAnalyzer {
    /// Creates a new security analyzer.
    pub fn new() -> Self {
        Self {
            findings: Vec::new(),
        }
    }

    /// Analyzes a parsed document for security issues.
    pub fn analyze(&mut self, store: &IrStore, root_id: docir_core::NodeId) -> AnalysisResult {
        self.findings.clear();

        // Run visitor analysis
        let mut walker = PreOrderWalker::new(store, root_id);
        if let Err(err) = walker.walk(self) {
            self.findings.push(ThreatIndicator {
                indicator_type: ThreatIndicatorType::SuspiciousApiCall,
                severity: ThreatLevel::Low,
                description: format!("Security walk incomplete: {err}"),
                location: None,
                node_id: None,
            });
        }

        // Get document security info
        let security_info = if let Some(IRNode::Document(doc)) = store.get(root_id) {
            Some(doc.security.clone())
        } else {
            None
        };

        // Build result
        AnalysisResult {
            threat_level: self.calculate_threat_level(&security_info),
            findings: std::mem::take(&mut self.findings),
            has_macros: security_info.as_ref().is_some_and(|s| s.has_macros()),
            has_ole_objects: security_info.as_ref().is_some_and(|s| s.has_ole_objects()),
            has_external_refs: security_info
                .as_ref()
                .is_some_and(|s| s.has_external_references()),
            has_dde: security_info.as_ref().is_some_and(|s| s.has_dde_fields()),
            has_xlm_macros: security_info.as_ref().is_some_and(|s| s.has_xlm_macros()),
        }
    }

    /// Calculate overall threat level.
    fn calculate_threat_level(&self, security_info: &Option<SecurityInfo>) -> ThreatLevel {
        let mut level = ThreatLevel::None;

        // Check security info
        if let Some(info) = security_info {
            if info.threat_level > level {
                level = info.threat_level;
            }
        }

        // Check findings
        for finding in &self.findings {
            if finding.severity > level {
                level = finding.severity;
            }
        }

        level
    }
}

impl Default for SecurityAnalyzer {
    fn default() -> Self {
        Self::new()
    }
}

impl IrVisitor for SecurityAnalyzer {
    fn visit_macro_project(
        &mut self,
        project: &docir_core::security::MacroProject,
    ) -> VisitorResult<VisitControl> {
        if project.has_auto_exec {
            self.findings.push(make_indicator(
                ThreatIndicatorType::AutoExecMacro,
                ThreatLevel::Critical,
                format!(
                    "VBA macro with auto-execute: {:?}",
                    project.auto_exec_procedures
                ),
                project.span.as_ref().map(|s| s.file_path.clone()),
                Some(project.id),
            ));
        }

        Ok(VisitControl::Continue)
    }

    fn visit_macro_module(
        &mut self,
        module: &docir_core::security::MacroModule,
    ) -> VisitorResult<VisitControl> {
        for call in &module.suspicious_calls {
            self.findings.push(make_indicator(
                ThreatIndicatorType::SuspiciousApiCall,
                ThreatLevel::High,
                format!(
                    "Suspicious API call in module '{}': {} ({:?})",
                    module.name, call.name, call.category
                ),
                module.span.as_ref().map(|s| s.file_path.clone()),
                Some(module.id),
            ));
        }

        Ok(VisitControl::Continue)
    }

    fn visit_ole_object(
        &mut self,
        ole: &docir_core::security::OleObject,
    ) -> VisitorResult<VisitControl> {
        let severity = if ole.is_linked {
            ThreatLevel::High
        } else {
            ThreatLevel::Medium
        };

        self.findings.push(make_indicator(
            ThreatIndicatorType::OleObject,
            severity,
            format!(
                "OLE object: {} ({} bytes)",
                ole.prog_id.as_deref().unwrap_or("unknown"),
                ole.size_bytes
            ),
            ole.span.as_ref().map(|s| s.file_path.clone()),
            Some(ole.id),
        ));

        Ok(VisitControl::Continue)
    }

    fn visit_external_ref(
        &mut self,
        ext_ref: &docir_core::security::ExternalReference,
    ) -> VisitorResult<VisitControl> {
        use docir_core::security::ExternalRefType;

        let (indicator_type, severity) = match ext_ref.ref_type {
            ExternalRefType::AttachedTemplate => {
                (ThreatIndicatorType::ExternalTemplate, ThreatLevel::High)
            }
            ExternalRefType::Hyperlink => {
                if ext_ref.is_remote() {
                    (ThreatIndicatorType::SuspiciousLink, ThreatLevel::Low)
                } else {
                    return Ok(VisitControl::Continue);
                }
            }
            ExternalRefType::Image if ext_ref.is_remote() => {
                (ThreatIndicatorType::RemoteResource, ThreatLevel::Medium)
            }
            ExternalRefType::OleLink => (ThreatIndicatorType::OleObject, ThreatLevel::High),
            _ => return Ok(VisitControl::Continue),
        };

        self.findings.push(make_indicator(
            indicator_type,
            severity,
            format!(
                "External reference ({:?}): {}",
                ext_ref.ref_type, ext_ref.target
            ),
            ext_ref.span.as_ref().map(|s| s.file_path.clone()),
            Some(ext_ref.id),
        ));

        Ok(VisitControl::Continue)
    }

    fn visit_activex_control(
        &mut self,
        control: &docir_core::security::ActiveXControl,
    ) -> VisitorResult<VisitControl> {
        self.findings.push(make_indicator(
            ThreatIndicatorType::ActiveXControl,
            ThreatLevel::High,
            format!(
                "ActiveX control: {} ({})",
                control.name.as_deref().unwrap_or("unknown"),
                control.clsid.as_deref().unwrap_or("unknown")
            ),
            control.span.as_ref().map(|s| s.file_path.clone()),
            Some(control.id),
        ));
        Ok(VisitControl::Continue)
    }
}

/// Result of security analysis.
#[derive(Debug, Clone)]
pub struct AnalysisResult {
    /// Overall threat level.
    pub threat_level: ThreatLevel,
    /// Specific findings.
    pub findings: Vec<ThreatIndicator>,
    /// Does the document contain VBA macros?
    pub has_macros: bool,
    /// Does the document contain OLE objects?
    pub has_ole_objects: bool,
    /// Does the document contain external references?
    pub has_external_refs: bool,
    /// Does the document contain DDE fields?
    pub has_dde: bool,
    /// Does the document contain XLM macros?
    pub has_xlm_macros: bool,
}

impl AnalysisResult {
    /// Returns true if the document has any security concerns.
    pub fn has_concerns(&self) -> bool {
        self.threat_level != ThreatLevel::None
    }

    /// Formats the result as a human-readable report.
    pub fn format_report(&self) -> String {
        let mut report = String::new();

        report.push_str(&format!("Threat Level: {}\n\n", self.threat_level));

        report.push_str("Security Features Detected:\n");
        report.push_str(&format!(
            "  - VBA Macros: {}\n",
            if self.has_macros { "YES" } else { "No" }
        ));
        report.push_str(&format!(
            "  - OLE Objects: {}\n",
            if self.has_ole_objects { "YES" } else { "No" }
        ));
        report.push_str(&format!(
            "  - External References: {}\n",
            if self.has_external_refs { "YES" } else { "No" }
        ));
        report.push_str(&format!(
            "  - DDE Fields: {}\n",
            if self.has_dde { "YES" } else { "No" }
        ));
        report.push_str(&format!(
            "  - XLM Macros: {}\n",
            if self.has_xlm_macros { "YES" } else { "No" }
        ));

        if !self.findings.is_empty() {
            report.push_str(&format!("\nFindings ({}):\n", self.findings.len()));
            for (i, finding) in self.findings.iter().enumerate() {
                report.push_str(&format!(
                    "  {}. [{}] {:?}: {}\n",
                    i + 1,
                    finding.severity,
                    finding.indicator_type,
                    finding.description
                ));
                if let Some(loc) = &finding.location {
                    report.push_str(&format!("     Location: {}\n", loc));
                }
            }
        }

        report
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::{Document, IRNode};
    use docir_core::security::{
        ActiveXControl, ExternalRefType, ExternalReference, MacroModule, MacroModuleType,
        MacroProject, OleObject, SuspiciousCall, SuspiciousCallCategory, ThreatIndicatorType,
        ThreatLevel,
    };
    use docir_core::types::{DocumentFormat, SourceSpan};
    use docir_core::visitor::IrStore;

    fn build_full_security_fixture() -> (IrStore, docir_core::NodeId) {
        let mut store = IrStore::new();
        let mut doc = Document::new(DocumentFormat::WordProcessing);
        doc.security.threat_level = ThreatLevel::Low;

        let mut project = MacroProject::new();
        project.has_auto_exec = true;
        project.auto_exec_procedures = vec!["AutoOpen".to_string()];
        project.span = Some(SourceSpan::new("vba/project.bin"));

        let mut module = MacroModule::new("Module1", MacroModuleType::Standard);
        module.suspicious_calls.push(SuspiciousCall {
            name: "Shell".to_string(),
            category: SuspiciousCallCategory::ShellExecution,
            line: Some(8),
        });
        module.span = Some(SourceSpan::new("vba/module1.bas"));
        project.modules.push(module.id);

        let mut ole = OleObject::new();
        ole.is_linked = true;
        ole.prog_id = Some("Package".to_string());
        ole.size_bytes = 2048;
        ole.span = Some(SourceSpan::new("word/embeddings/ole1.bin"));

        let mut ext_template = ExternalReference::new(
            ExternalRefType::AttachedTemplate,
            "http://evil/template.dotm",
        );
        ext_template.span = Some(SourceSpan::new("word/_rels/document.xml.rels"));
        let ext_hyperlink =
            ExternalReference::new(ExternalRefType::Hyperlink, "https://example.test/phish");
        let ext_local_link =
            ExternalReference::new(ExternalRefType::Hyperlink, "file:///tmp/report");
        let ext_remote_image =
            ExternalReference::new(ExternalRefType::Image, "https://cdn.example.test/a.png");
        let ext_ole_link =
            ExternalReference::new(ExternalRefType::OleLink, "file:///tmp/linked.ole");

        let mut activex = ActiveXControl::new();
        activex.name = Some("CommandButton1".to_string());
        activex.clsid = Some("{ABC}".to_string());
        activex.span = Some(SourceSpan::new("word/activeX/activeX1.xml"));

        doc.security.macro_project = Some(project.id);
        doc.security.ole_objects.push(ole.id);
        doc.security.external_refs.extend([
            ext_template.id,
            ext_hyperlink.id,
            ext_local_link.id,
            ext_remote_image.id,
            ext_ole_link.id,
        ]);
        doc.security.activex_controls.push(activex.id);

        let root_id = doc.id;
        store.insert(IRNode::Document(doc));
        store.insert(IRNode::MacroProject(project));
        store.insert(IRNode::MacroModule(module));
        store.insert(IRNode::OleObject(ole));
        store.insert(IRNode::ExternalReference(ext_template));
        store.insert(IRNode::ExternalReference(ext_hyperlink));
        store.insert(IRNode::ExternalReference(ext_local_link));
        store.insert(IRNode::ExternalReference(ext_remote_image));
        store.insert(IRNode::ExternalReference(ext_ole_link));
        store.insert(IRNode::ActiveXControl(activex));
        (store, root_id)
    }

    #[test]
    fn analyze_collects_findings_and_escalates_to_critical() {
        let (store, root_id) = build_full_security_fixture();
        let mut analyzer = SecurityAnalyzer::new();
        let result = analyzer.analyze(&store, root_id);

        assert_eq!(result.threat_level, ThreatLevel::Critical);
        assert!(result.has_concerns());
        assert!(result.has_macros);
        assert!(result.has_ole_objects);
        assert!(result.has_external_refs);
        assert!(!result.has_dde);
        assert!(!result.has_xlm_macros);
        assert_eq!(result.findings.len(), 8);
    }

    #[test]
    fn analyze_marks_expected_indicator_types() {
        let (store, root_id) = build_full_security_fixture();
        let mut analyzer = SecurityAnalyzer::new();
        let result = analyzer.analyze(&store, root_id);

        assert!(result
            .findings
            .iter()
            .any(|f| f.indicator_type == ThreatIndicatorType::AutoExecMacro));
        assert!(result
            .findings
            .iter()
            .any(|f| f.indicator_type == ThreatIndicatorType::SuspiciousApiCall));
        assert!(result
            .findings
            .iter()
            .any(|f| f.indicator_type == ThreatIndicatorType::ExternalTemplate));
        assert!(result
            .findings
            .iter()
            .any(|f| f.indicator_type == ThreatIndicatorType::SuspiciousLink));
        assert!(result
            .findings
            .iter()
            .any(|f| f.indicator_type == ThreatIndicatorType::RemoteResource));
        assert!(result
            .findings
            .iter()
            .any(|f| f.indicator_type == ThreatIndicatorType::ActiveXControl));
    }

    #[test]
    fn analyze_report_includes_summary_and_location() {
        let (store, root_id) = build_full_security_fixture();
        let mut analyzer = SecurityAnalyzer::new();
        let result = analyzer.analyze(&store, root_id);
        let report = result.format_report();
        assert!(report.contains("Threat Level: CRITICAL"));
        assert!(report.contains("VBA Macros: YES"));
        assert!(report.contains("OLE Objects: YES"));
        assert!(report.contains("External References: YES"));
        assert!(report.contains("Findings (8):"));
        assert!(report.contains("Location: vba/project.bin"));
    }

    #[test]
    fn analyze_non_document_root_uses_findings_without_security_flags() {
        let mut store = IrStore::new();
        let mut ole = OleObject::new();
        ole.is_linked = false;
        ole.size_bytes = 64;
        let root_id = ole.id;
        store.insert(IRNode::OleObject(ole));

        let mut analyzer = SecurityAnalyzer::new();
        let result = analyzer.analyze(&store, root_id);

        assert_eq!(result.threat_level, ThreatLevel::Medium);
        assert_eq!(result.findings.len(), 1);
        assert!(!result.has_macros);
        assert!(!result.has_ole_objects);
        assert!(!result.has_external_refs);
        assert!(!result.has_dde);
        assert!(!result.has_xlm_macros);
    }
}
