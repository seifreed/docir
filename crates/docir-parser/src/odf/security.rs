use crate::diagnostics::push_entry;
use crate::security_utils::parse_dde_formula;
use crate::zip_handler::PackageReader;
use docir_core::ir::{DiagnosticEntry, DiagnosticSeverity, Diagnostics, Document, IRNode};
use docir_core::security::{DdeField, ExternalRefType, ExternalReference, OleObject};
use docir_core::types::SourceSpan;
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

pub(crate) struct OdfFormulaScan {
    pub(crate) dde_fields: Vec<DdeField>,
    pub(crate) external_refs: Vec<ExternalReference>,
    pub(crate) diagnostics: Vec<DiagnosticEntry>,
}

impl Default for OdfFormulaScan {
    fn default() -> Self {
        Self {
            dde_fields: Vec::new(),
            external_refs: Vec::new(),
            diagnostics: Vec::new(),
        }
    }
}

fn visit_start_or_empty(xml: &str, mut on_element: impl FnMut(&BytesStart<'_>) -> bool) {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if on_element(&e) {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

pub(crate) fn scan_external_links(xml: &str, location: &str) -> Vec<ExternalReference> {
    let mut refs = Vec::new();
    visit_start_or_empty(xml, |e| {
        if let Some(href) = super::attr_value(e, b"xlink:href") {
            let ref_type = match e.name().as_ref() {
                b"draw:image" => ExternalRefType::Image,
                b"text:a" => ExternalRefType::Hyperlink,
                b"draw:object" | b"draw:object-ole" => ExternalRefType::OleLink,
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
        match e.name().as_ref() {
            b"draw:object" | b"draw:object-ole" => {
                if let Some(href) = super::attr_value(e, b"xlink:href") {
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
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if let Some(formula_attr) = super::attr_value(&e, b"table:formula") {
                    process_formula(&formula_attr, &mut scan, &mut unsupported, &mut has_array);
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
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
    content_xml: Option<&str>,
    styles_xml: Option<&str>,
    settings_xml: Option<&str>,
    file_names: &[String],
    zip: &mut impl PackageReader,
    store: &mut IrStore,
    doc: &mut Document,
    diagnostics: &mut Diagnostics,
) {
    let mut formula_scan = OdfFormulaScan::default();
    if let Some(xml) = content_xml {
        formula_scan = scan_odf_formula_security(xml);
        diagnostics
            .entries
            .extend(formula_scan.diagnostics.drain(..));
        diagnostics.entries.extend(scan_odf_protection(xml));
        diagnostics.entries.extend(scan_odf_advanced_features(xml));
    }

    let mut external_refs = Vec::new();
    for (xml, location) in [
        (content_xml, "content.xml"),
        (styles_xml, "styles.xml"),
        (settings_xml, "settings.xml"),
    ] {
        if let Some(xml) = xml {
            external_refs.extend(scan_external_links(xml, location));
        }
    }
    external_refs.extend(formula_scan.external_refs.drain(..));

    let mut ole_objects = Vec::new();
    if let Some(xml) = content_xml {
        let (oles, ole_links) = scan_odf_objects(xml);
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
    doc.security
        .dde_fields
        .extend(formula_scan.dde_fields.drain(..));
}

pub(crate) fn scan_odf_protection(xml: &str) -> Vec<DiagnosticEntry> {
    let mut entries = Vec::new();
    let mut protected = false;
    visit_start_or_empty(xml, |e| {
        if let Some(value) = super::attr_value(e, b"table:protected")
            .or_else(|| super::attr_value(e, b"text:protected"))
        {
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
        match e.name().as_ref() {
            b"table:conditional-format" => {
                if let Some(condition) = super::attr_value(e, b"table:condition") {
                    if super::parse_odf_condition_operator(&condition).is_none() {
                        conditional_advanced = true;
                    }
                }
            }
            b"table:pivot-table" | b"table:data-pilot-table" => {
                if super::attr_value(e, b"table:target-range-address").is_some() {
                    pivot_advanced = true;
                }
            }
            b"draw:object" | b"draw:object-ole" => odp_advanced = true,
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
            if matches!(lookahead.peek(), Some('(')) {
                if !ident.is_empty() {
                    functions.push(ident.to_ascii_uppercase());
                }
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
            while let Some(c) = chars.next() {
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
