use crate::error::ParseError;
use crate::xml_utils::{
    XmlScanControl, local_name, reader_from_str, scan_xml_events, try_attr_value,
    try_attr_value_by_suffix, xml_error,
};
use docir_core::ir::IRNode;
use docir_core::security::{MacroModule, MacroModuleType, MacroProject};
use docir_core::visitor::IrStore;
use quick_xml::events::Event;

use super::manifest::OdfManifestEntry;

pub(crate) fn build_odf_macro_project(
    manifest_entries: &[OdfManifestEntry],
    content_xml: &Option<String>,
    styles_xml: &Option<String>,
    settings_xml: &Option<String>,
    file_names: &[String],
    store: &mut IrStore,
) -> Result<Option<MacroProject>, ParseError> {
    let mut module_paths = Vec::new();
    for entry in manifest_entries {
        if let Some(media) = entry.media_type.as_deref()
            && (media.contains("script") || media.contains("basic"))
        {
            module_paths.push(entry.path.clone());
        }
    }
    for name in file_names {
        if name.starts_with("Scripts/") || name.starts_with("Basic/") || name.ends_with(".bas") {
            module_paths.push(name.clone());
        }
    }

    if let Some(xml) = content_xml.as_deref() {
        module_paths.extend(scan_script_links(xml, "content.xml")?);
    }
    if let Some(xml) = styles_xml.as_deref() {
        module_paths.extend(scan_script_links(xml, "styles.xml")?);
    }
    if let Some(xml) = settings_xml.as_deref() {
        module_paths.extend(scan_script_links(xml, "settings.xml")?);
    }

    module_paths.sort();
    module_paths.dedup();

    if module_paths.is_empty() {
        return Ok(None);
    }

    let mut project = MacroProject::new();
    project.name = Some("ODF Scripts".to_string());

    for path in module_paths {
        let module = MacroModule::new(path.clone(), MacroModuleType::Standard);
        let module_id = module.id;
        store.insert(IRNode::MacroModule(module));
        project.modules.push(module_id);
    }

    Ok(Some(project))
}

pub(crate) fn scan_script_links(xml: &str, source: &str) -> Result<Vec<String>, ParseError> {
    let mut links = Vec::new();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();

    scan_xml_events(&mut reader, &mut buf, source, |event| {
        match event {
            Event::Start(e) | Event::Empty(e) => {
                if local_name(e.name().as_ref()) == b"script"
                    && let Some(href) = try_attr_value_by_suffix(&e, &[b":href"], source)?
                {
                    links.push(href);
                }
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok(links)
}

pub(crate) fn parse_odf_signatures(
    xml: &str,
    source: &str,
) -> Result<Vec<docir_core::ir::DigitalSignature>, ParseError> {
    let mut sigs = Vec::new();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let mut current: Option<docir_core::ir::DigitalSignature> = None;

    scan_xml_events(&mut reader, &mut buf, source, |event| {
        match event {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"Signature" => current = Some(docir_core::ir::DigitalSignature::new()),
                b"SignatureMethod" => {
                    if let Some(sig) = current.as_mut() {
                        sig.signature_method = try_attr_value(&e, b"Algorithm", source)?;
                    }
                }
                b"DigestMethod" => {
                    if let Some(sig) = current.as_mut()
                        && let Some(alg) = try_attr_value(&e, b"Algorithm", source)?
                    {
                        sig.digest_methods.push(alg);
                    }
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"SignatureMethod" => {
                    if let Some(sig) = current.as_mut() {
                        sig.signature_method = try_attr_value(&e, b"Algorithm", source)?;
                    }
                }
                b"DigestMethod" => {
                    if let Some(sig) = current.as_mut()
                        && let Some(alg) = try_attr_value(&e, b"Algorithm", source)?
                    {
                        sig.digest_methods.push(alg);
                    }
                }
                _ => {}
            },
            Event::Text(e) => {
                if let Some(sig) = current.as_mut() {
                    let text =
                        crate::xml_utils::decoded_text(&e).map_err(|err| xml_error(source, err))?;
                    if sig.signer.is_none() && text.contains("CN=") {
                        sig.signer = Some(text);
                    }
                }
            }
            Event::GeneralRef(e) => {
                if let Some(sig) = current.as_mut() {
                    let text = crate::xml_utils::decoded_general_ref(&e)
                        .map_err(|err| xml_error(source, err))?;
                    if sig.signer.is_none() && text.contains("CN=") {
                        sig.signer = Some(text);
                    }
                }
            }
            Event::End(e) => {
                if local_name(e.name().as_ref()) == b"Signature"
                    && let Some(sig) = current.take()
                {
                    sigs.push(sig);
                }
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok(sigs)
}

#[cfg(test)]
mod tests {
    use super::{parse_odf_signatures, scan_script_links};
    use crate::error::ParseError;

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

        assert_eq!(
            scan_script_links(xml, "content.xml").expect("valid script links XML"),
            vec!["Scripts/macro.py"]
        );
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

        let signatures = parse_odf_signatures(xml, "META-INF/documentsignatures.xml")
            .expect("valid signatures XML");
        assert_eq!(signatures.len(), 1);
        assert_eq!(
            signatures[0].signature_method.as_deref(),
            Some("rsa-sha256")
        );
        assert_eq!(signatures[0].digest_methods, vec!["sha256"]);
        assert_eq!(signatures[0].signer.as_deref(), Some("CN=Tester"));
    }

    #[test]
    fn parse_odf_signatures_reports_malformed_attributes() {
        let xml = r#"
            <sig:Signatures xmlns:sig="http://www.w3.org/2000/09/xmldsig#">
              <sig:Signature>
                <sig:SignedInfo>
                  <sig:SignatureMethod Algorithm="rsa-sha256" Algorithm="rsa-sha512"/>
                </sig:SignedInfo>
              </sig:Signature>
            </sig:Signatures>
        "#;

        let err = parse_odf_signatures(xml, "META-INF/documentsignatures.xml")
            .expect_err("malformed ODF signature attributes must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "META-INF/documentsignatures.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }
}
