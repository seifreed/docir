use crate::diagnostics::push_entry;
use crate::security_utils::parse_dde_formula;
use crate::zip_handler::SecureZipReader;
use docir_core::ir::{DiagnosticEntry, DiagnosticSeverity};
use docir_core::security::{DdeField, ExternalRefType, ExternalReference, OleObject};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::{Read, Seek};

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

pub(crate) fn scan_external_links(xml: &str, location: &str) -> Vec<ExternalReference> {
    let mut refs = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if let Some(href) = super::attr_value(&e, b"xlink:href") {
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
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    refs
}

pub(crate) fn scan_odf_objects(xml: &str) -> (Vec<OleObject>, Vec<ExternalReference>) {
    let mut oles = Vec::new();
    let mut refs = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"draw:object" | b"draw:object-ole" => {
                    if let Some(href) = super::attr_value(&e, b"xlink:href") {
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
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    (oles, refs)
}

pub(crate) fn scan_embedded_objects(
    file_names: &[String],
    zip: &mut SecureZipReader<impl Read + Seek>,
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
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref().starts_with(b"table:filter") {
                    let target = super::attr_value(&e, b"table:target-range-address")
                        .or_else(|| super::attr_value(&e, b"table:condition"))
                        .unwrap_or_else(|| "unknown".to_string());
                    out.push(target);
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

pub(crate) fn scan_odf_formula_security(xml: &str) -> OdfFormulaScan {
    let mut scan = OdfFormulaScan::default();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let supported = ["SUM", "AVERAGE", "MIN", "MAX", "COUNT"];
    let mut unsupported: Vec<String> = Vec::new();
    let mut has_array = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if let Some(formula_attr) = super::attr_value(&e, b"table:formula") {
                    let formula_raw = unescape_xml_value(&formula_attr);
                    let formula = super::strip_odf_formula_prefix(&formula_raw)
                        .trim()
                        .to_string();
                    if formula.contains('{') || formula.contains('}') {
                        has_array = true;
                    }
                    if let Some(dde) =
                        parse_dde_formula(&formula, SourceSpan::new("content.xml"), false)
                    {
                        scan.dde_fields.push(dde);
                    }
                    let func_names = extract_formula_functions(&formula);
                    for name in func_names {
                        if supported.contains(&name.as_str()) {
                            continue;
                        }
                        if name == "DDE" || name == "DDEAUTO" {
                            continue;
                        }
                        if !unsupported.contains(&name) {
                            unsupported.push(name);
                        }
                    }
                    let lower = formula.to_ascii_lowercase();
                    let ref_type = if lower.contains("hyperlink(") {
                        ExternalRefType::Hyperlink
                    } else {
                        ExternalRefType::DataConnection
                    };
                    for target in extract_quoted_strings(&formula) {
                        let target_lower = target.to_ascii_lowercase();
                        if target_lower.contains("://")
                            || target_lower.starts_with("file:")
                            || target_lower.starts_with("smb:")
                            || target_lower.starts_with("ftp:")
                            || target_lower.starts_with("mailto:")
                        {
                            let mut ext = ExternalReference::new(ref_type, target);
                            ext.span = Some(SourceSpan::new("content.xml"));
                            scan.external_refs.push(ext);
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    if !unsupported.is_empty() {
        push_entry(
            &mut scan.diagnostics,
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
            &mut scan.diagnostics,
            DiagnosticSeverity::Info,
            "ODF_FORMULA_ARRAY",
            "ODF array formula detected (not fully evaluated)".to_string(),
            Some("content.xml"),
        );
    }
    scan
}

pub(crate) fn scan_odf_protection(xml: &str) -> Vec<DiagnosticEntry> {
    let mut entries = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut protected = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if let Some(value) = super::attr_value(&e, b"table:protected")
                    .or_else(|| super::attr_value(&e, b"text:protected"))
                {
                    if value == "true" {
                        protected = true;
                        break;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
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
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut conditional_advanced = false;
    let mut pivot_advanced = false;
    let mut odp_advanced = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:conditional-format" => {
                    if let Some(condition) = super::attr_value(&e, b"table:condition") {
                        if super::parse_odf_condition_operator(&condition).is_none() {
                            conditional_advanced = true;
                        }
                    }
                }
                b"table:pivot-table" | b"table:data-pilot-table" => {
                    if super::attr_value(&e, b"table:target-range-address").is_some() {
                        pivot_advanced = true;
                    }
                }
                b"draw:object" | b"draw:object-ole" => odp_advanced = true,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
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
