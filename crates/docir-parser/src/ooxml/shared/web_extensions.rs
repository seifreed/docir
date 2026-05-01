use crate::error::ParseError;
use crate::xml_utils::{attr_each, local_name, read_event};
use docir_core::ir::{WebExtension, WebExtensionProperty, WebExtensionTaskpane};
use docir_core::types::SourceSpan;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

type AttrList = Vec<(Vec<u8>, String)>;

/// Public API entrypoint: parse_web_extension.
pub fn parse_web_extension(xml: &str, path: &str) -> Result<WebExtension, ParseError> {
    parse_web_extension_impl(xml, path)
}

fn parse_web_extension_impl(xml: &str, path: &str) -> Result<WebExtension, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut ext = WebExtension::new();
    ext.span = Some(SourceSpan::new(path));

    let mut buf = Vec::new();
    loop {
        match read_event(&mut reader, &mut buf, path)? {
            Event::Start(e) | Event::Empty(e) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                let attrs = collect_local_attrs(&e);

                match local {
                    b"webextension" => {
                        if let Some(value) = find_attr(&attrs, &[b"id", b"rId", b"rid"]) {
                            ext.extension_id = Some(value.to_string());
                        }
                    }
                    b"storeReference" | b"storereference" => {
                        apply_web_extension_store_reference(&mut ext, &attrs);
                    }
                    b"reference" => {
                        apply_web_extension_reference(&mut ext, &attrs);
                    }
                    b"property" => {
                        if let Some((name, value)) = parse_web_extension_property(&attrs) {
                            ext.properties.push(WebExtensionProperty { name, value });
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(ext)
}

/// Public API entrypoint: parse_web_extension_taskpanes.
pub fn parse_web_extension_taskpanes(
    xml: &str,
    path: &str,
) -> Result<Vec<WebExtensionTaskpane>, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut panes: Vec<WebExtensionTaskpane> = Vec::new();
    let mut current: Option<WebExtensionTaskpane> = None;

    let mut buf = Vec::new();
    loop {
        match read_event(&mut reader, &mut buf, path)? {
            Event::Start(e) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if local == b"taskpane" {
                    current = Some(new_taskpane(path, &e));
                } else if local == b"webextensionref" {
                    if let Some(pane) = current.as_mut() {
                        apply_webextension_ref_attrs(pane, &collect_local_attrs(&e));
                    }
                }
            }
            Event::Empty(e) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if local == b"taskpane" {
                    panes.push(new_taskpane(path, &e));
                } else if local == b"webextensionref" {
                    if let Some(pane) = current.as_mut() {
                        apply_webextension_ref_attrs(pane, &collect_local_attrs(&e));
                    }
                }
            }
            Event::End(e) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if local == b"taskpane" {
                    if let Some(pane) = current.take() {
                        panes.push(pane);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(panes)
}

fn new_taskpane(path: &str, e: &BytesStart<'_>) -> WebExtensionTaskpane {
    let mut pane = WebExtensionTaskpane::new();
    pane.span = Some(SourceSpan::new(path));
    apply_taskpane_attrs(&mut pane, &collect_local_attrs(e));
    pane
}

fn apply_taskpane_attrs(pane: &mut WebExtensionTaskpane, attrs: &AttrList) {
    for (key, val) in attrs {
        match key.as_slice() {
            b"dockState" | b"dockstate" => pane.dock_state = Some(val.clone()),
            b"visibility" => {
                pane.visibility = Some(val.as_bytes() == b"1" || val.eq_ignore_ascii_case("true"));
            }
            b"width" => pane.width = val.parse::<u32>().ok(),
            b"height" => pane.height = val.parse::<u32>().ok(),
            b"row" => pane.row = val.parse::<u32>().ok(),
            b"column" => pane.column = val.parse::<u32>().ok(),
            _ => {}
        }
    }
}

fn apply_web_extension_store_reference(ext: &mut WebExtension, attrs: &AttrList) {
    for (key, val) in attrs {
        match key.as_slice() {
            b"store" => ext.store = Some(val.clone()),
            b"storeType" | b"storetype" => ext.store_type = Some(val.clone()),
            b"id" => ext.store_id = Some(val.clone()),
            b"version" => ext.version = Some(val.clone()),
            _ => {}
        }
    }
}

fn apply_web_extension_reference(ext: &mut WebExtension, attrs: &AttrList) {
    for (key, val) in attrs {
        match key.as_slice() {
            b"id" => ext.reference_id = Some(val.clone()),
            b"version" => ext.reference_version = Some(val.clone()),
            b"store" => ext.store = Some(val.clone()),
            b"storeType" | b"storetype" => ext.store_type = Some(val.clone()),
            _ => {}
        }
    }
}

fn apply_webextension_ref_attrs(pane: &mut WebExtensionTaskpane, attrs: &AttrList) {
    if let Some(value) = find_attr(attrs, &[b"id", b"rid", b"rId"]) {
        pane.web_extension_ref = Some(value.to_string());
    }
}

fn parse_web_extension_property(attrs: &AttrList) -> Option<(String, String)> {
    let name = find_attr(attrs, &[b"name"])?;
    let value = find_attr(attrs, &[b"value", b"val"])?;
    Some((name.to_string(), value.to_string()))
}

fn collect_local_attrs(element: &BytesStart<'_>) -> AttrList {
    let mut attrs = Vec::new();
    attr_each(element, |key, value| {
        attrs.push((key.to_vec(), String::from_utf8_lossy(value).to_string()));
    });
    attrs
}

fn find_attr<'a>(attrs: &'a AttrList, keys: &[&[u8]]) -> Option<&'a str> {
    attrs.iter().find_map(|(key, val)| {
        if keys.contains(&key.as_slice()) {
            Some(val.as_str())
        } else {
            None
        }
    })
}
