use crate::xml_utils::{attr_value, attr_value_by_suffix, local_name};
use docir_core::ir::IRNode;
use docir_core::security::{MacroModule, MacroModuleType, MacroProject};
use docir_core::visitor::IrStore;
use quick_xml::events::Event;
use quick_xml::Reader;

use super::manifest::OdfManifestEntry;

pub(crate) fn build_odf_macro_project(
    manifest_entries: &[OdfManifestEntry],
    content_xml: &Option<String>,
    styles_xml: &Option<String>,
    settings_xml: &Option<String>,
    file_names: &[String],
    store: &mut IrStore,
) -> Option<MacroProject> {
    let mut module_paths = Vec::new();
    for entry in manifest_entries {
        if let Some(media) = entry.media_type.as_deref() {
            if media.contains("script") || media.contains("basic") {
                module_paths.push(entry.path.clone());
            }
        }
    }
    for name in file_names {
        if name.starts_with("Scripts/") || name.starts_with("Basic/") || name.ends_with(".bas") {
            module_paths.push(name.clone());
        }
    }

    if let Some(xml) = content_xml.as_deref() {
        module_paths.extend(scan_script_links(xml));
    }
    if let Some(xml) = styles_xml.as_deref() {
        module_paths.extend(scan_script_links(xml));
    }
    if let Some(xml) = settings_xml.as_deref() {
        module_paths.extend(scan_script_links(xml));
    }

    module_paths.sort();
    module_paths.dedup();

    if module_paths.is_empty() {
        return None;
    }

    let mut project = MacroProject::new();
    project.name = Some("ODF Scripts".to_string());

    for path in module_paths {
        let module = MacroModule::new(path.clone(), MacroModuleType::Standard);
        let module_id = module.id;
        store.insert(IRNode::MacroModule(module));
        project.modules.push(module_id);
    }

    Some(project)
}

pub(crate) fn scan_script_links(xml: &str) -> Vec<String> {
    let mut links = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if local_name(e.name().as_ref()) == b"script" {
                    if let Some(href) = attr_value_by_suffix(&e, &[b":href"]) {
                        links.push(href);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    links
}

pub(crate) fn parse_odf_signatures(xml: &str) -> Vec<docir_core::ir::DigitalSignature> {
    let mut sigs = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut current: Option<docir_core::ir::DigitalSignature> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"Signature" => current = Some(docir_core::ir::DigitalSignature::new()),
                b"SignatureMethod" => {
                    if let Some(sig) = current.as_mut() {
                        sig.signature_method = attr_value(&e, b"Algorithm");
                    }
                }
                b"DigestMethod" => {
                    if let Some(sig) = current.as_mut() {
                        if let Some(alg) = attr_value(&e, b"Algorithm") {
                            sig.digest_methods.push(alg);
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                b"SignatureMethod" => {
                    if let Some(sig) = current.as_mut() {
                        sig.signature_method = attr_value(&e, b"Algorithm");
                    }
                }
                b"DigestMethod" => {
                    if let Some(sig) = current.as_mut() {
                        if let Some(alg) = attr_value(&e, b"Algorithm") {
                            sig.digest_methods.push(alg);
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if let Some(sig) = current.as_mut() {
                    let text = match e.unescape() {
                        Ok(t) => t.to_string(),
                        Err(_) => String::new(),
                    };
                    if sig.signer.is_none() && text.contains("CN=") {
                        sig.signer = Some(text);
                    }
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"Signature" {
                    if let Some(sig) = current.take() {
                        sigs.push(sig);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    sigs
}

#[cfg(test)]
mod tests {
    use super::{parse_odf_signatures, scan_script_links};

    #[test]
    fn scan_script_links_accepts_alternate_namespace_prefixes() {
        let xml = r#"
            <office:document-content
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:scr="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
              xmlns:lnk="http://www.w3.org/1999/xlink">
              <office:scripts>
                <scr:script lnk:href="Scripts/macro.py"/>
              </office:scripts>
            </office:document-content>
        "#;

        assert_eq!(scan_script_links(xml), vec!["Scripts/macro.py"]);
    }

    #[test]
    fn parse_odf_signatures_accepts_alternate_namespace_prefixes() {
        let xml = r#"
            <sig:Signatures xmlns:sig="http://www.w3.org/2000/09/xmldsig#">
              <sig:Signature>
                <sig:SignedInfo>
                  <sig:SignatureMethod Algorithm="rsa-sha256"/>
                  <sig:Reference>
                    <sig:DigestMethod Algorithm="sha256"/>
                  </sig:Reference>
                </sig:SignedInfo>
                <sig:KeyInfo>
                  <sig:X509SubjectName>CN=Tester</sig:X509SubjectName>
                </sig:KeyInfo>
              </sig:Signature>
            </sig:Signatures>
        "#;

        let signatures = parse_odf_signatures(xml);
        assert_eq!(signatures.len(), 1);
        assert_eq!(
            signatures[0].signature_method.as_deref(),
            Some("rsa-sha256")
        );
        assert_eq!(signatures[0].digest_methods, vec!["sha256"]);
        assert_eq!(signatures[0].signer.as_deref(), Some("CN=Tester"));
    }
}
