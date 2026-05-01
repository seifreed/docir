//! OOXML Part Registry for coverage tracking.

use crate::registry_utils::matches_pattern;
use docir_core::types::DocumentFormat;

mod presentation;
mod spreadsheet;
mod word;

use presentation::PRESENTATION_PARTS;
use spreadsheet::SPREADSHEET_PARTS;
use word::WORD_PARTS;

#[derive(Debug, Clone)]
pub struct PartSpec {
    pub format: DocumentFormat,
    pub pattern: &'static str,
    pub content_type: Option<&'static str>,
    pub expected_parser: &'static str,
}

impl PartSpec {
    pub fn matches(&self, path: &str, content_type_value: Option<&str>) -> bool {
        if let Some(ct) = self.content_type {
            if let Some(actual) = content_type_value {
                if actual != ct {
                    return false;
                }
            }
        }
        matches_pattern(path, self.pattern)
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct PartSpecEntry {
    pattern: &'static str,
    content_type: Option<&'static str>,
    expected_parser: &'static str,
}

fn build_registry(format: DocumentFormat, entries: &[PartSpecEntry]) -> Vec<PartSpec> {
    entries
        .iter()
        .map(|entry| PartSpec {
            format,
            pattern: entry.pattern,
            content_type: entry.content_type,
            expected_parser: entry.expected_parser,
        })
        .collect()
}

/// Public API entrypoint: registry_for.
pub fn registry_for(format: DocumentFormat) -> Vec<PartSpec> {
    match format {
        DocumentFormat::WordProcessing => build_registry(format, WORD_PARTS),
        DocumentFormat::Spreadsheet => build_registry(format, SPREADSHEET_PARTS),
        DocumentFormat::Presentation => build_registry(format, PRESENTATION_PARTS),
        DocumentFormat::OdfText
        | DocumentFormat::OdfSpreadsheet
        | DocumentFormat::OdfPresentation
        | DocumentFormat::Hwp
        | DocumentFormat::Hwpx
        | DocumentFormat::Rtf => Vec::new(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_registry_includes_macro_projects() {
        let xlsx = registry_for(DocumentFormat::Spreadsheet);
        assert!(xlsx.iter().any(|p| p.pattern == "xl/vbaProject.bin"));

        let pptx = registry_for(DocumentFormat::Presentation);
        assert!(pptx.iter().any(|p| p.pattern == "ppt/vbaProject.bin"));
    }

    #[test]
    fn test_registry_includes_package_parts_and_rels() {
        let docx = registry_for(DocumentFormat::WordProcessing);
        assert!(docx.iter().any(|p| p.pattern == "_rels/.rels"));
        assert!(docx.iter().any(|p| p.pattern == "word/_rels/*.rels"));
        assert!(docx.iter().any(|p| p.pattern == "docProps/core.xml"));
        assert!(docx.iter().any(|p| p.pattern == "docProps/app.xml"));
        assert!(docx.iter().any(|p| p.pattern == "docProps/custom.xml"));
        assert!(docx.iter().any(|p| p.pattern == "customXml/*.xml"));
        assert!(docx.iter().any(|p| p.pattern == "customXml/itemProps*.xml"));
    }

    #[test]
    fn test_registry_includes_signature_and_customxml_props() {
        let docx = registry_for(DocumentFormat::WordProcessing);
        assert!(docx.iter().any(|p| p.pattern == "_xmlsignatures/*.xml"));
        assert!(docx
            .iter()
            .any(|p| p.pattern == "_xmlsignatures/origin.sigs"));
        assert!(docx.iter().any(|p| p.pattern == "customXml/itemProps*.xml"));
    }
}
