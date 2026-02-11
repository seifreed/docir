//! Security analyzer for IR documents.

use crate::make_indicator;
use docir_core::ir::{Document, IRNode};
use docir_core::security::{
    SecurityInfo, ThreatIndicator, ThreatIndicatorType, ThreatLevel, AUTO_EXEC_PROCEDURES,
    DANGEROUS_XLM_FUNCTIONS, SUSPICIOUS_VBA_CALLS,
};
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
        let _ = walker.walk(self);

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
            has_macros: security_info
                .as_ref()
                .map_or(false, |s| s.macro_project.is_some()),
            has_ole_objects: security_info
                .as_ref()
                .map_or(false, |s| !s.ole_objects.is_empty()),
            has_external_refs: security_info
                .as_ref()
                .map_or(false, |s| !s.external_refs.is_empty()),
            has_dde: security_info
                .as_ref()
                .map_or(false, |s| !s.dde_fields.is_empty()),
            has_xlm_macros: security_info
                .as_ref()
                .map_or(false, |s| !s.xlm_macros.is_empty()),
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
