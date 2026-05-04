use crate::error::ParseError;
use crate::xml_utils::{scan_xml_events, xml_error, XmlScanControl};
use docir_core::ir::{SharedStringItem, SharedStringTable};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;
use quick_xml::Reader;

pub(crate) fn parse_shared_strings_table(
    xml: &str,
) -> Result<(SharedStringTable, Vec<String>), ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut strings: Vec<String> = Vec::new();
    let mut table = SharedStringTable::new();
    table.span = Some(SourceSpan::new("xl/sharedStrings.xml"));

    let mut in_si = false;
    let mut in_t = false;
    let mut in_run = false;
    let mut current = String::new();
    let mut current_run = String::new();
    let mut runs: Vec<String> = Vec::new();

    scan_xml_events(&mut reader, &mut buf, "xl/sharedStrings.xml", |event| {
        match event {
            Event::Start(e) => match e.name().as_ref() {
                b"si" => {
                    in_si = true;
                    current.clear();
                    current_run.clear();
                    runs.clear();
                }
                b"r" if in_si => {
                    in_run = true;
                    current_run.clear();
                }
                b"t" if in_si => {
                    in_t = true;
                }
                _ => {}
            },
            Event::Text(e) => {
                if in_si && in_t {
                    let text = e
                        .unescape()
                        .map_err(|err| xml_error("xl/sharedStrings.xml", err))?;
                    current.push_str(&text);
                    if in_run {
                        current_run.push_str(&text);
                    }
                }
            }
            Event::End(e) => match e.name().as_ref() {
                b"t" => in_t = false,
                b"r" => {
                    if in_run {
                        runs.push(current_run.clone());
                        in_run = false;
                        current_run.clear();
                    }
                }
                b"si" => {
                    in_si = false;
                    strings.push(current.clone());
                    table.items.push(SharedStringItem {
                        text: current.clone(),
                        runs: runs.clone(),
                    });
                }
                _ => {}
            },
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok((table, strings))
}

#[cfg(test)]
mod tests {
    use super::parse_shared_strings_table;

    #[test]
    fn test_parse_shared_strings() {
        let xml = r#"
        <sst>
          <si><t>Hello</t></si>
          <si><r><t>Foo</t></r><r><t>Bar</t></r></si>
        </sst>
        "#;
        let (table, strings) = parse_shared_strings_table(xml).expect("shared strings");
        assert_eq!(strings, vec!["Hello", "FooBar"]);
        assert_eq!(table.items.len(), 2);
        assert_eq!(
            table.items[1].runs,
            vec!["Foo".to_string(), "Bar".to_string()]
        );
    }
}
