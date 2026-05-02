use crate::xml_utils::attr_value;
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
                if e.name().as_ref() == b"script:script" {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
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
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"ds:Signature" => current = Some(docir_core::ir::DigitalSignature::new()),
                b"ds:SignatureMethod" => {
                    if let Some(sig) = current.as_mut() {
                        sig.signature_method = attr_value(&e, b"Algorithm");
                    }
                }
                b"ds:DigestMethod" => {
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
                if e.name().as_ref() == b"ds:Signature" {
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
