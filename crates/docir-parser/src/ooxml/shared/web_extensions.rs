use crate::error::ParseError;
use crate::xml_utils::{local_name, read_event};
use docir_core::ir::{WebExtension, WebExtensionProperty, WebExtensionTaskpane};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;
use quick_xml::Reader;

pub fn parse_web_extension(xml: &str, path: &str) -> Result<WebExtension, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut ext = WebExtension::new();
    ext.span = Some(SourceSpan::new(path));

    let mut buf = Vec::new();
    loop {
        match read_event(&mut reader, &mut buf, path)? {
            Event::Start(e) | Event::Empty(e) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                match local {
                    b"webextension" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"id" || key == b"rId" || key == b"rid" {
                                ext.extension_id = Some(val);
                            }
                        }
                    }
                    b"storeReference" | b"storereference" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key {
                                b"store" => ext.store = Some(val),
                                b"storeType" | b"storetype" => ext.store_type = Some(val),
                                b"id" => ext.store_id = Some(val),
                                b"version" => ext.version = Some(val),
                                _ => {}
                            }
                        }
                    }
                    b"reference" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key {
                                b"id" => ext.reference_id = Some(val),
                                b"version" => ext.reference_version = Some(val),
                                b"store" => ext.store = Some(val),
                                b"storeType" | b"storetype" => ext.store_type = Some(val),
                                _ => {}
                            }
                        }
                    }
                    b"property" => {
                        let mut name = None;
                        let mut value = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key {
                                b"name" => name = Some(val),
                                b"value" | b"val" => value = Some(val),
                                _ => {}
                            }
                        }
                        if let (Some(name), Some(value)) = (name, value) {
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
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"taskpane" {
                    let mut pane = WebExtensionTaskpane::new();
                    pane.span = Some(SourceSpan::new(path));
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key {
                            b"dockState" | b"dockstate" => pane.dock_state = Some(val),
                            b"visibility" => {
                                let v = val.eq_ignore_ascii_case("true") || val == "1";
                                pane.visibility = Some(v);
                            }
                            b"width" => pane.width = val.parse::<u32>().ok(),
                            b"height" => pane.height = val.parse::<u32>().ok(),
                            b"row" => pane.row = val.parse::<u32>().ok(),
                            b"column" => pane.column = val.parse::<u32>().ok(),
                            _ => {}
                        }
                    }
                    current = Some(pane);
                } else if local == b"webextensionref" {
                    if let Some(pane) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"id" || key == b"rid" || key == b"rId" {
                                pane.web_extension_ref = Some(val);
                            }
                        }
                    }
                }
            }
            Event::Empty(e) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"taskpane" {
                    let mut pane = WebExtensionTaskpane::new();
                    pane.span = Some(SourceSpan::new(path));
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key {
                            b"dockState" | b"dockstate" => pane.dock_state = Some(val),
                            b"visibility" => {
                                let v = val.eq_ignore_ascii_case("true") || val == "1";
                                pane.visibility = Some(v);
                            }
                            b"width" => pane.width = val.parse::<u32>().ok(),
                            b"height" => pane.height = val.parse::<u32>().ok(),
                            b"row" => pane.row = val.parse::<u32>().ok(),
                            b"column" => pane.column = val.parse::<u32>().ok(),
                            _ => {}
                        }
                    }
                    panes.push(pane);
                } else if local == b"webextensionref" {
                    if let Some(pane) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"id" || key == b"rid" || key == b"rId" {
                                pane.web_extension_ref = Some(val);
                            }
                        }
                    }
                }
            }
            Event::End(e) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
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
