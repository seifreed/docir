use super::helpers::parse_odf_condition_operator;
use crate::diagnostics::push_entry;
use crate::security_scan::OdfXmlInputs;
use crate::security_utils::parse_dde_formula;
use crate::xml_utils::{attr_value_by_suffix, local_name, scan_xml_events, XmlScanControl};
use crate::zip_handler::PackageReader;
use docir_core::ir::{DiagnosticEntry, DiagnosticSeverity, Diagnostics, Document, IRNode};
use docir_core::security::{DdeField, ExternalRefType, ExternalReference, OleObject};
use docir_core::types::SourceSpan;
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

#[derive(Default)]
pub(crate) struct OdfFormulaScan {
    pub(crate) dde_fields: Vec<DdeField>,
    pub(crate) external_refs: Vec<ExternalReference>,
    pub(crate) diagnostics: Vec<DiagnosticEntry>,
}

fn visit_start_or_empty(xml: &str, mut on_element: impl FnMut(&BytesStart<'_>) -> bool) {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let _ = scan_xml_events(&mut reader, &mut buf, "content.xml", |event| {
        match event {
            Event::Start(e) | Event::Empty(e) => {
                if on_element(&e) {
                    return Ok(XmlScanControl::Break);
                }
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    });
}

pub(crate) fn scan_external_links(xml: &str, location: &str) -> Vec<ExternalReference> {
    let mut refs = Vec::new();
    visit_start_or_empty(xml, |e| {
        if let Some(href) = attr_value_by_suffix(e, &[b":href"]) {
            let ref_type = match local_name(e.name().as_ref()) {
                b"image" => ExternalRefType::Image,
                b"a" => ExternalRefType::Hyperlink,
                b"object" | b"object-ole" => ExternalRefType::OleLink,
                _ => ExternalRefType::Other,
            };
            let mut ext = ExternalReference::new(ref_type, href);
            ext.span = Some(SourceSpan::new(location));
            refs.push(ext);
        }
        false
    });
    refs
}

pub(crate) fn scan_odf_objects(xml: &str) -> (Vec<OleObject>, Vec<ExternalReference>) {
    let mut oles = Vec::new();
    let mut refs = Vec::new();
    visit_start_or_empty(xml, |e| {
        match local_name(e.name().as_ref()) {
            b"object" | b"object-ole" => {
                if let Some(href) = attr_value_by_suffix(e, &[b":href"]) {
                    let mut ole = OleObject::new();
                    ole.is_linked = href.starts_with("http://") || href.starts_with("https://");
                    ole.link_target = Some(href.clone());
                    ole.size_bytes = 0;
                    oles.push(ole);
                    let mut ext = ExternalReference::new(ExternalRefType::OleLink, href);
                    ext.span = Some(SourceSpan::new("content.xml"));
                    refs.push(ext);
                }
            }
            _ => {}
        }
        false
    });
    (oles, refs)
}

pub(crate) fn scan_embedded_objects(
    file_names: &[String],
    zip: &mut impl PackageReader,
) -> Vec<OleObject> {
    let mut oles = Vec::new();
    for path in file_names {
        if path.starts_with("Object ")
            || path.starts_with("ObjectReplacements/")
            || path.starts_with("Objects/")
        {
            let size_bytes = zip.file_size(path).unwrap_or(0);
            let mut ole = OleObject::new();
            ole.name = Some(path.clone());
            ole.size_bytes = size_bytes;
            ole.is_linked = false;
            oles.push(ole);
        }
    }
    oles
}

pub(crate) fn scan_odf_filters(xml: &str) -> Vec<String> {
    let mut out = Vec::new();
    visit_start_or_empty(xml, |e| {
        if e.name().as_ref().starts_with(b"table:filter") {
            let target = super::attr_value(e, b"table:target-range-address")
                .or_else(|| super::attr_value(e, b"table:condition"))
                .unwrap_or_else(|| "unknown".to_string());
            out.push(target);
        }
        false
    });
    out
}

pub(crate) fn scan_odf_formula_security(xml: &str) -> OdfFormulaScan {
    let mut scan = OdfFormulaScan::default();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut unsupported: Vec<String> = Vec::new();
    let mut has_array = false;
    let _ = scan_xml_events(&mut reader, &mut buf, "content.xml", |event| {
        match event {
            Event::Start(e) | Event::Empty(e) => {
                if let Some(formula_attr) = attr_value_by_suffix(&e, &[b":formula"]) {
                    process_formula(&formula_attr, &mut scan, &mut unsupported, &mut has_array);
                }
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    });
    push_formula_diagnostics(&mut scan.diagnostics, &unsupported, has_array);
    scan
}

fn process_formula(
    formula_attr: &str,
    scan: &mut OdfFormulaScan,
    unsupported: &mut Vec<String>,
    has_array: &mut bool,
) {
    let formula = normalize_formula(formula_attr);
    if is_array_formula(&formula) {
        *has_array = true;
    }

    if let Some(dde) = parse_dde_formula(&formula, SourceSpan::new("content.xml"), false) {
        scan.dde_fields.push(dde);
    }

    collect_unsupported_functions(&formula, unsupported);
    scan.external_refs
        .extend(extract_formula_external_refs(&formula));
}

fn normalize_formula(formula_attr: &str) -> String {
    let formula_raw = unescape_xml_value(formula_attr);
    super::strip_odf_formula_prefix(&formula_raw)
        .trim()
        .to_string()
}

fn is_array_formula(formula: &str) -> bool {
    formula.contains('{') || formula.contains('}')
}

fn collect_unsupported_functions(formula: &str, unsupported: &mut Vec<String>) {
    let supported = ["SUM", "AVERAGE", "MIN", "MAX", "COUNT"];
    for name in extract_formula_functions(formula) {
        if supported.contains(&name.as_str()) || name == "DDE" || name == "DDEAUTO" {
            continue;
        }
        if !unsupported.contains(&name) {
            unsupported.push(name);
        }
    }
}

fn extract_formula_external_refs(formula: &str) -> Vec<ExternalReference> {
    let lower = formula.to_ascii_lowercase();
    let ref_type = if lower.contains("hyperlink(") {
        ExternalRefType::Hyperlink
    } else {
        ExternalRefType::DataConnection
    };

    extract_quoted_strings(formula)
        .into_iter()
        .filter(|target| is_external_target(target))
        .map(|target| {
            let mut ext = ExternalReference::new(ref_type, target);
            ext.span = Some(SourceSpan::new("content.xml"));
            ext
        })
        .collect()
}

fn is_external_target(target: &str) -> bool {
    let lower = target.to_ascii_lowercase();
    lower.contains("://")
        || lower.starts_with("file:")
        || lower.starts_with("smb:")
        || lower.starts_with("ftp:")
        || lower.starts_with("mailto:")
}

fn push_formula_diagnostics(
    diagnostics: &mut Vec<DiagnosticEntry>,
    unsupported: &[String],
    has_array: bool,
) {
    if !unsupported.is_empty() {
        push_entry(
            diagnostics,
            DiagnosticSeverity::Info,
            "ODF_FORMULA_UNSUPPORTED_FUNCTION",
            format!(
                "Unsupported ODF formula functions detected: {}",
                unsupported.join(", ")
            ),
            Some("content.xml"),
        );
    }
    if has_array {
        push_entry(
            diagnostics,
            DiagnosticSeverity::Info,
            "ODF_FORMULA_ARRAY",
            "ODF array formula detected (not fully evaluated)".to_string(),
            Some("content.xml"),
        );
    }
}

pub(crate) fn scan_odf_security(
    xml: OdfXmlInputs<'_>,
    file_names: &[String],
    zip: &mut impl PackageReader,
    store: &mut IrStore,
    doc: &mut Document,
    diagnostics: &mut Diagnostics,
) {
    let mut formula_scan = OdfFormulaScan::default();
    if let Some(content_xml) = xml.content_xml {
        formula_scan = scan_odf_formula_security(content_xml);
        diagnostics.entries.append(&mut formula_scan.diagnostics);
        diagnostics.entries.extend(scan_odf_protection(content_xml));
        diagnostics
            .entries
            .extend(scan_odf_advanced_features(content_xml));
    }

    let mut external_refs = Vec::new();
    for (xml, location) in [
        (xml.content_xml, "content.xml"),
        (xml.styles_xml, "styles.xml"),
        (xml.settings_xml, "settings.xml"),
    ] {
        if let Some(xml) = xml {
            external_refs.extend(scan_external_links(xml, location));
        }
    }
    external_refs.append(&mut formula_scan.external_refs);

    let mut ole_objects = Vec::new();
    if let Some(content_xml) = xml.content_xml {
        let (oles, ole_links) = scan_odf_objects(content_xml);
        ole_objects.extend(oles);
        external_refs.extend(ole_links);
    }
    ole_objects.extend(scan_embedded_objects(file_names, zip));

    for ext in external_refs {
        store.insert(IRNode::ExternalReference(ext));
    }
    for ole in ole_objects {
        store.insert(IRNode::OleObject(ole));
    }
    doc.security.dde_fields.append(&mut formula_scan.dde_fields);
}

pub(crate) fn scan_odf_protection(xml: &str) -> Vec<DiagnosticEntry> {
    let mut entries = Vec::new();
    let mut protected = false;
    visit_start_or_empty(xml, |e| {
        if let Some(value) = attr_value_by_suffix(e, &[b":protected"]) {
            if value == "true" {
                protected = true;
                return true;
            }
        }
        false
    });
    if protected {
        push_entry(
            &mut entries,
            DiagnosticSeverity::Info,
            "ODF_PROTECTED_CONTENT",
            "ODF protected content detected".to_string(),
            Some("content.xml"),
        );
    }
    entries
}

pub(crate) fn scan_odf_advanced_features(xml: &str) -> Vec<DiagnosticEntry> {
    let mut entries = Vec::new();
    let mut conditional_advanced = false;
    let mut pivot_advanced = false;
    let mut odp_advanced = false;
    visit_start_or_empty(xml, |e| {
        match local_name(e.name().as_ref()) {
            b"conditional-format" => {
                if let Some(condition) = attr_value_by_suffix(e, &[b":condition"]) {
                    if parse_odf_condition_operator(&condition).is_none() {
                        conditional_advanced = true;
                    }
                }
            }
            b"pivot-table" | b"data-pilot-table" => {
                if attr_value_by_suffix(e, &[b":target-range-address"]).is_some() {
                    pivot_advanced = true;
                }
            }
            b"object" | b"object-ole" => odp_advanced = true,
            _ => {}
        }
        false
    });
    if conditional_advanced {
        push_entry(
            &mut entries,
            DiagnosticSeverity::Info,
            "ODF_CONDITIONAL_ADVANCED",
            "Advanced conditional formatting detected".to_string(),
            Some("content.xml"),
        );
    }
    if pivot_advanced {
        push_entry(
            &mut entries,
            DiagnosticSeverity::Info,
            "ODF_PIVOT_ADVANCED",
            "Advanced pivot table features detected".to_string(),
            Some("content.xml"),
        );
    }
    if odp_advanced {
        push_entry(
            &mut entries,
            DiagnosticSeverity::Info,
            "ODF_ODP_OBJECTS",
            "ODP embedded objects detected".to_string(),
            Some("content.xml"),
        );
    }
    entries
}

fn unescape_xml_value(value: &str) -> String {
    match quick_xml::escape::unescape(value) {
        Ok(cow) => cow.into_owned(),
        Err(_) => value.to_string(),
    }
}

fn extract_formula_functions(formula: &str) -> Vec<String> {
    let mut functions = Vec::new();
    let mut chars = formula.chars().peekable();
    while let Some(&ch) = chars.peek() {
        if ch.is_ascii_alphabetic() || ch == '_' {
            let mut ident = String::new();
            while let Some(&c) = chars.peek() {
                if c.is_ascii_alphanumeric() || c == '_' || c == '.' || c == '$' {
                    ident.push(c);
                    chars.next();
                } else {
                    break;
                }
            }
            let mut lookahead = chars.clone();
            while let Some(&c) = lookahead.peek() {
                if c.is_whitespace() {
                    lookahead.next();
                } else {
                    break;
                }
            }
            if matches!(lookahead.peek(), Some('(')) && !ident.is_empty() {
                functions.push(ident.to_ascii_uppercase());
            }
        } else {
            chars.next();
        }
    }
    functions
}

fn extract_quoted_strings(input: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut chars = input.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '"' || ch == '\'' {
            let quote = ch;
            let mut value = String::new();
            for c in chars.by_ref() {
                if c == quote {
                    break;
                }
                value.push(c);
            }
            if !value.is_empty() {
                out.push(value);
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::{scan_external_links, scan_odf_formula_security, scan_odf_objects};
    use docir_core::security::ExternalRefType;

    #[test]
    fn scan_external_links_accepts_alternate_namespace_prefixes() {
        let xml = r#"
            <office:document-content
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:txt="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
              xmlns:dr="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
              xmlns:lnk="http://www.w3.org/1999/xlink">
              <txt:a lnk:href="https://example.test/link">Link</txt:a>
              <dr:image lnk:href="Pictures/pic.png"/>
              <dr:object-ole lnk:href="https://example.test/object.bin"/>
            </office:document-content>
        "#;

        let refs = scan_external_links(xml, "content.xml");
        let types: Vec<_> = refs.iter().map(|r| r.ref_type).collect();

        assert_eq!(refs.len(), 3);
        assert!(types.contains(&ExternalRefType::Hyperlink));
        assert!(types.contains(&ExternalRefType::Image));
        assert!(types.contains(&ExternalRefType::OleLink));
    }

    #[test]
    fn scan_odf_objects_accepts_alternate_namespace_prefixes() {
        let xml = r#"
            <office:document-content
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:dr="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
              xmlns:lnk="http://www.w3.org/1999/xlink">
              <dr:object-ole lnk:href="https://example.test/object.bin"/>
            </office:document-content>
        "#;

        let (oles, refs) = scan_odf_objects(xml);
        assert_eq!(oles.len(), 1);
        assert_eq!(
            oles[0].link_target.as_deref(),
            Some("https://example.test/object.bin")
        );
        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0].ref_type, ExternalRefType::OleLink);
    }

    #[test]
    fn scan_odf_formula_security_accepts_alternate_formula_prefix() {
        let xml = r#"
            <office:document-content
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:tbl="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
              <tbl:table-cell tbl:formula='of:=DDE("cmd";"/c calc";"A1")'/>
            </office:document-content>
        "#;

        let scan = scan_odf_formula_security(xml);
        assert_eq!(scan.dde_fields.len(), 1);
        assert_eq!(scan.dde_fields[0].application, "cmd");
    }
}
