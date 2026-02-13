use super::*;

struct HwpxSecurityScan {
    external_refs: Vec<ExternalReference>,
    macro_modules: Vec<NodeId>,
    has_autoexec: bool,
    encrypted_flag: bool,
    protected_flag: bool,
}

impl HwpxSecurityScan {
    fn new() -> Self {
        Self {
            external_refs: Vec::new(),
            macro_modules: Vec::new(),
            has_autoexec: false,
            encrypted_flag: false,
            protected_flag: false,
        }
    }

    fn mark_flags_from_text(&mut self, text: &str) {
        let lower = text.to_ascii_lowercase();
        if lower.contains("encrypt") {
            self.encrypted_flag = true;
        }
        if lower.contains("protect") || lower.contains("password") {
            self.protected_flag = true;
        }
    }
}

pub(super) fn scan_hwpx_security<R: Read + Seek>(
    file_names: &[String],
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
    doc: &mut Document,
) {
    let mut scan = HwpxSecurityScan::new();
    let mut diagnostics = Diagnostics::new();
    diagnostics.span = Some(SourceSpan::new("package"));

    for path in file_names {
        scan_hwpx_security_path(path, zip, store, &mut scan);
    }

    for ext in scan.external_refs {
        store.insert(IRNode::ExternalReference(ext));
    }
    if !scan.macro_modules.is_empty() {
        let mut project = MacroProject::new();
        project.name = Some("HWPX Scripts".to_string());
        project.modules = scan.macro_modules;
        project.has_auto_exec = scan.has_autoexec;
        project.span = Some(SourceSpan::new("package"));
        store.insert(IRNode::MacroProject(project));
    }
    if scan.encrypted_flag {
        push_warning(
            &mut diagnostics,
            "HWPX_ENCRYPTED",
            "HWPX encrypted content detected".to_string(),
            None,
        );
    }
    if scan.protected_flag {
        push_info(
            &mut diagnostics,
            "HWPX_PROTECTED",
            "HWPX protected content detected".to_string(),
            None,
        );
    }
    attach_diagnostics_if_any(store, doc, diagnostics);
}

fn scan_hwpx_security_path<R: Read + Seek>(
    path: &str,
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
    scan: &mut HwpxSecurityScan,
) {
    let lower = path.to_ascii_lowercase();
    scan.mark_flags_from_text(&lower);
    scan_hwpx_binary_ole(path, &lower, zip, store);
    scan_hwpx_script(path, &lower, zip, store, scan);
    if path.ends_with(".xml") {
        scan_hwpx_xml(path, zip, store, scan);
    }
}

fn scan_hwpx_binary_ole<R: Read + Seek>(
    path: &str,
    lower: &str,
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
) {
    if !lower.starts_with("bindata/") {
        return;
    }
    if !(lower.contains("ole") || lower.contains("object")) {
        return;
    }
    let mut ole = OleObject::new();
    ole.name = Some(path.to_string());
    ole.size_bytes = zip.file_size(path).unwrap_or(0);
    store.insert(IRNode::OleObject(ole));
}

fn scan_hwpx_script<R: Read + Seek>(
    path: &str,
    lower: &str,
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
    scan: &mut HwpxSecurityScan,
) {
    let is_script = lower.starts_with("scripts/")
        || lower.contains("/scripts/")
        || lower.starts_with("macros/")
        || lower.contains("/macros/")
        || lower.ends_with(".js")
        || lower.ends_with(".vbs")
        || lower.ends_with(".wsf")
        || lower.ends_with(".sct")
        || lower.ends_with(".py");
    if !is_script {
        return;
    }

    if let Ok(data) = zip.read_file(path) {
        let source = String::from_utf8_lossy(&data).to_string();
        let mut module = MacroModule::new(path.to_string(), MacroModuleType::Standard);
        module.source_code = Some(source.clone());
        module.span = Some(SourceSpan::new(path));
        let id = module.id;
        store.insert(IRNode::MacroModule(module));
        scan.macro_modules.push(id);

        let lower_source = source.to_ascii_lowercase();
        if lower.contains("auto")
            || lower_source.contains("autoexec")
            || lower_source.contains("auto_open")
            || lower_source.contains("onopen")
        {
            scan.has_autoexec = true;
        }
    }
}

fn scan_hwpx_xml<R: Read + Seek>(
    path: &str,
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
    scan: &mut HwpxSecurityScan,
) {
    if let Ok(xml) = zip.read_file_string(path) {
        scan.external_refs
            .extend(scan_hwpx_external_refs(&xml, path));
        scan.mark_flags_from_text(&xml);
        if xml.to_ascii_lowercase().contains("ole") {
            let mut ole = OleObject::new();
            ole.name = Some(path.to_string());
            ole.size_bytes = xml.len() as u64;
            store.insert(IRNode::OleObject(ole));
        }
    }
}

fn scan_hwpx_external_refs(xml: &str, source: &str) -> Vec<ExternalReference> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut refs = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    if key.ends_with(b"href") || key.ends_with(b"src") || key.ends_with(b"link") {
                        if let Ok(value) = attr.unescape_value() {
                            let target = value.to_string();
                            if target.is_empty() {
                                continue;
                            }
                            let mut ext =
                                ExternalReference::new(ExternalRefType::Hyperlink, target);
                            ext.span = Some(SourceSpan::new(source));
                            refs.push(ext);
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
    refs
}
