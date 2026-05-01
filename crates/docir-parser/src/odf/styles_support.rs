#[path = "styles_support_headers_footers.rs"]
mod styles_support_headers_footers;
#[path = "styles_support_styles.rs"]
mod styles_support_styles;
pub(crate) use styles_support_headers_footers::parse_odf_headers_footers;
pub(crate) use styles_support_styles::{
    merge_styles, parse_master_pages, parse_page_layouts, parse_styles,
};

#[cfg(test)]
use crate::error::ParseError;
#[cfg(test)]
use crate::parser::ParserConfig;
#[cfg(test)]
use crate::IrStore;
#[cfg(test)]
use docir_core::ir::{IRNode, Style, StyleSet, StyleType};

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_styles_supports_default_family_fallback_and_text_props() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
  <style:default-style style:family="paragraph">
    <style:text-properties fo:font-family="Fira Sans" fo:font-size="11pt"/>
    <style:paragraph-properties fo:text-align="center"/>
  </style:default-style>
</office:document-styles>
"#;

        let styles = parse_styles(xml).expect("expected style set");
        assert_eq!(styles.styles.len(), 1);

        let style = &styles.styles[0];
        assert_eq!(style.style_id, "default:paragraph");
        assert_eq!(style.style_type, StyleType::Paragraph);
        assert!(style.is_default);
        assert_eq!(
            style
                .run_props
                .as_ref()
                .and_then(|p| p.font_family.as_deref()),
            Some("Fira Sans")
        );
        assert_eq!(style.run_props.as_ref().and_then(|p| p.font_size), Some(11));
        assert_eq!(
            style
                .paragraph_props
                .as_ref()
                .and_then(|p| p.alignment)
                .map(|alignment| format!("{alignment:?}")),
            Some("Center".to_string())
        );
    }

    #[test]
    fn parse_styles_returns_none_for_malformed_xml() {
        let malformed = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">"#;
        assert!(parse_styles(malformed).is_none());
    }

    #[test]
    fn parse_odf_headers_footers_reports_styles_xml_errors() {
        let malformed = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
  <style:header>
    <text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">broken
</office:document-styles>
"#;
        let mut store = IrStore::new();
        let err =
            parse_odf_headers_footers(malformed, &mut store, &ParserConfig::default()).unwrap_err();

        match err {
            ParseError::Xml { file, message } => {
                assert!(file == "styles.xml" || file == "content.xml");
                assert!(!message.is_empty());
            }
            other => panic!("expected styles.xml parse error, got {:?}", other),
        }
    }

    #[test]
    fn parse_styles_handles_family_mapping_defaults_and_font_units() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
  <style:style style:name="ListStyle" style:family="list"/>
  <style:style style:name="TextDefault" style:family="text" style:default="true">
    <style:text-properties fo:font-size="13px" fo:font-weight="bold" fo:font-style="italic"/>
  </style:style>
  <style:style style:name="OtherStyle" style:family="custom">
    <style:text-properties fo:font-size="bad-unit"/>
  </style:style>
</office:document-styles>
"#;

        let styles = parse_styles(xml).expect("expected styles");
        assert_eq!(styles.styles.len(), 3);

        let list_style = styles
            .styles
            .iter()
            .find(|s| s.style_id == "ListStyle")
            .expect("missing list style");
        assert_eq!(list_style.style_type, StyleType::Numbering);
        assert!(!list_style.is_default);

        let text_default = styles
            .styles
            .iter()
            .find(|s| s.style_id == "TextDefault")
            .expect("missing text style");
        assert_eq!(text_default.style_type, StyleType::Character);
        assert!(text_default.is_default);
        let run_props = text_default.run_props.as_ref().expect("missing run props");
        assert_eq!(run_props.font_size, Some(13));
        assert_eq!(run_props.bold, Some(true));
        assert_eq!(run_props.italic, Some(true));

        let other_style = styles
            .styles
            .iter()
            .find(|s| s.style_id == "OtherStyle")
            .expect("missing custom style");
        assert_eq!(other_style.style_type, StyleType::Other);
        assert_eq!(
            other_style
                .run_props
                .as_ref()
                .and_then(|props| props.font_size),
            None
        );
    }

    #[test]
    fn merge_styles_keeps_first_style_when_ids_overlap() {
        let mut existing = StyleSet::new();
        existing.styles.push(Style {
            style_id: "S1".to_string(),
            name: Some("existing".to_string()),
            style_type: StyleType::Paragraph,
            based_on: None,
            next: None,
            is_default: false,
            run_props: None,
            paragraph_props: None,
            table_props: None,
        });

        let mut incoming = StyleSet::new();
        incoming.styles.push(Style {
            style_id: "S1".to_string(),
            name: Some("incoming".to_string()),
            style_type: StyleType::Character,
            based_on: None,
            next: None,
            is_default: false,
            run_props: None,
            paragraph_props: None,
            table_props: None,
        });
        incoming.styles.push(Style {
            style_id: "S2".to_string(),
            name: Some("new".to_string()),
            style_type: StyleType::Table,
            based_on: None,
            next: None,
            is_default: false,
            run_props: None,
            paragraph_props: None,
            table_props: None,
        });

        merge_styles(&mut existing, &mut incoming);

        assert_eq!(existing.styles.len(), 2);
        assert_eq!(
            existing.styles[0].name.as_deref(),
            Some("existing"),
            "existing style must win on duplicate IDs"
        );
        assert!(existing.styles.iter().any(|s| s.style_id == "S2"));
        assert!(incoming.styles.is_empty());
    }

    #[test]
    fn parse_master_pages_and_layouts_skip_entries_without_names() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
  <style:master-page/>
  <style:master-page style:name="Standard"/>
  <style:page-layout/>
  <style:page-layout style:name="pm1"/>
</office:document-styles>
"#;

        assert_eq!(parse_master_pages(xml), vec!["Standard".to_string()]);
        assert_eq!(parse_page_layouts(xml), vec!["pm1".to_string()]);
    }

    #[test]
    fn parse_odf_headers_footers_handles_left_variants_lists_and_empty_headings() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <style:master-page style:name="Standard">
    <style:header-left>
      <text:h text:outline-level="2"/>
      <text:list text:style-name="L1">
        <text:list-item><text:p>Item one</text:p></text:list-item>
        <text:list-item><text:p>Item two</text:p></text:list-item>
      </text:list>
    </style:header-left>
    <style:footer-left>
      <text:p/>
    </style:footer-left>
  </style:master-page>
</office:document-styles>
"#;

        let mut store = IrStore::new();
        let (headers, footers) =
            parse_odf_headers_footers(xml, &mut store, &ParserConfig::default())
                .expect("parse headers/footers");
        assert_eq!(headers.len(), 1);
        assert_eq!(footers.len(), 1);

        let Some(IRNode::Header(header)) = store.get(headers[0]) else {
            panic!("expected header node");
        };
        assert_eq!(
            header.span.as_ref().map(|s| s.file_path.as_str()),
            Some("styles.xml")
        );
        assert_eq!(header.content.len(), 3);

        let Some(IRNode::Paragraph(heading)) = store.get(header.content[0]) else {
            panic!("expected heading paragraph");
        };
        assert_eq!(heading.properties.outline_level, Some(2));
        assert!(heading.properties.numbering.is_none());

        let Some(IRNode::Paragraph(list_item_1)) = store.get(header.content[1]) else {
            panic!("expected first list paragraph");
        };
        let Some(IRNode::Paragraph(list_item_2)) = store.get(header.content[2]) else {
            panic!("expected second list paragraph");
        };
        let numbering_1 = list_item_1
            .properties
            .numbering
            .as_ref()
            .expect("first list item should have numbering");
        let numbering_2 = list_item_2
            .properties
            .numbering
            .as_ref()
            .expect("second list item should have numbering");
        assert_eq!(numbering_1.level, 0);
        assert_eq!(numbering_2.level, 0);
        assert_eq!(numbering_1.num_id, numbering_2.num_id);

        let Some(IRNode::Footer(footer)) = store.get(footers[0]) else {
            panic!("expected footer node");
        };
        assert_eq!(
            footer.span.as_ref().map(|s| s.file_path.as_str()),
            Some("styles.xml")
        );
        assert_eq!(footer.content.len(), 1);
        let Some(IRNode::Paragraph(footer_para)) = store.get(footer.content[0]) else {
            panic!("expected footer paragraph");
        };
        assert!(footer_para.properties.numbering.is_none());
    }
}
