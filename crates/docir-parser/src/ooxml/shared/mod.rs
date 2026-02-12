//! Shared OOXML parts parsing (themes, relationships, custom XML, media).

mod charts;
mod custom_xml;
mod drawingml;
mod media;
mod people;
mod relationships;
mod signatures;
mod theme;
mod vml;
mod web_extensions;

pub use charts::parse_chart_data;
pub use custom_xml::parse_custom_xml_part;
pub use drawingml::parse_drawingml_part;
pub use media::{classify_media_type, legacy_extension_part};
pub use people::parse_people_part;
pub use relationships::build_relationship_graph;
pub use signatures::parse_signature;
pub use theme::parse_theme;
pub use vml::parse_vml_drawing;
pub use web_extensions::{parse_web_extension, parse_web_extension_taskpanes};

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::relationships::Relationships;
    use docir_core::ir::{RelationshipTargetMode, ShapeType};

    #[test]
    fn test_parse_theme_basic() {
        let xml = r#"
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
          <a:themeElements>
            <a:clrScheme name="Office">
              <a:dk1><a:srgbClr val="000000"/></a:dk1>
              <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
            </a:clrScheme>
            <a:fontScheme name="Office">
              <a:majorFont><a:latin typeface="Calibri"/></a:majorFont>
              <a:minorFont><a:latin typeface="Calibri Light"/></a:minorFont>
            </a:fontScheme>
          </a:themeElements>
        </a:theme>
        "#;
        let theme = parse_theme(xml, "theme/theme1.xml").expect("theme parse");
        assert_eq!(theme.name.as_deref(), Some("Office"));
        assert!(theme.colors.len() >= 2);
        assert_eq!(theme.fonts.major.as_deref(), Some("Calibri"));
    }

    #[test]
    fn test_parse_people_part() {
        let xml = r#"
        <ppl:people xmlns:ppl="http://schemas.openxmlformats.org/officeDocument/2006/relationships/people">
          <ppl:person ppl:id="p1" ppl:userId="user1" ppl:displayName="Alice" ppl:initials="A" />
          <ppl:person ppl:id="p2" ppl:userId="user2" ppl:displayName="Bob" />
        </ppl:people>
        "#;

        let people = parse_people_part(xml, "word/people.xml").expect("people");
        assert_eq!(people.people.len(), 2);
        assert_eq!(people.people[0].display_name.as_deref(), Some("Alice"));
        assert_eq!(people.people[1].display_name.as_deref(), Some("Bob"));
    }

    #[test]
    fn test_parse_web_extension() {
        let xml = r#"
        <we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11">
          <we:storeReference store="Store" storeType="OMEX" id="ext-1" version="1.0.0"/>
          <we:reference id="ref-1" version="1.0.0"/>
          <we:properties>
            <we:property name="foo" value="bar"/>
          </we:properties>
        </we:webextension>
        "#;

        let ext = parse_web_extension(xml, "word/webExtensions/webExtension1.xml").expect("ext");
        assert_eq!(ext.store.as_deref(), Some("Store"));
        assert_eq!(ext.store_type.as_deref(), Some("OMEX"));
        assert_eq!(ext.store_id.as_deref(), Some("ext-1"));
        assert_eq!(ext.reference_id.as_deref(), Some("ref-1"));
        assert_eq!(ext.properties.len(), 1);
    }

    #[test]
    fn test_parse_web_extension_taskpanes() {
        let xml = r#"
        <wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <wetp:taskpane dockState="right" visibility="true" width="300">
            <wetp:webextensionref r:id="rId1"/>
          </wetp:taskpane>
        </wetp:taskpanes>
        "#;

        let panes = parse_web_extension_taskpanes(xml, "word/webExtensions/taskpanes.xml")
            .expect("taskpanes");
        assert_eq!(panes.len(), 1);
        assert_eq!(panes[0].dock_state.as_deref(), Some("right"));
        assert_eq!(panes[0].web_extension_ref.as_deref(), Some("rId1"));
        assert_eq!(panes[0].width, Some(300));
    }

    #[test]
    fn test_parse_vml_drawing() {
        let xml = r##"
        <xml xmlns:v="urn:schemas-microsoft-com:vml"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <v:shape id="shape1" type="#_x0000_t75" style="width:100pt;height:50pt;">
            <v:imagedata r:id="rId1"/>
          </v:shape>
        </xml>
        "##;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels parse");
        let (drawing, shapes) = parse_vml_drawing(xml, "word/vmlDrawing1.vml", &rels).expect("vml");
        assert_eq!(drawing.path, "word/vmlDrawing1.vml");
        assert_eq!(shapes.len(), 1);
        assert_eq!(shapes[0].name.as_deref(), Some("shape1"));
        assert_eq!(shapes[0].image_target.as_deref(), Some("media/image1.png"));
    }

    #[test]
    fn test_parse_drawingml_part() {
        let xml = r#"
        <w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                   xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                   xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                   xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <wp:inline>
            <wp:docPr id="1" name="Image 1"/>
            <a:graphic>
              <a:graphicData>
                <a:blip r:embed="rId1"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
          <wp:inline>
            <wp:docPr id="2" name="Chart 1"/>
            <a:graphic>
              <a:graphicData>
                <c:chart r:id="rId2"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
          <wp:inline>
            <wp:docPr id="3" name="SmartArt 1"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <dgm:relIds r:dm="rId3" r:lo="rId4"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>
          <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="diagrams/layout1.xml"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels parse");
        let (_part, shapes) =
            parse_drawingml_part(xml, "word/drawings/drawing1.xml", &rels).expect("drawingml");
        assert_eq!(shapes.len(), 3);
        assert_eq!(shapes[0].name.as_deref(), Some("Image 1"));
        assert_eq!(
            shapes[0].media_target.as_deref(),
            Some("word/media/image1.png")
        );
        assert_eq!(shapes[1].shape_type, ShapeType::Chart);
        assert_eq!(
            shapes[1].media_target.as_deref(),
            Some("word/charts/chart1.xml")
        );
        assert_eq!(shapes[2].name.as_deref(), Some("SmartArt 1"));
        assert_eq!(shapes[2].shape_type, ShapeType::Custom);
        assert!(shapes[2]
            .related_targets
            .contains(&"word/diagrams/data1.xml".to_string()));
        assert!(shapes[2]
            .related_targets
            .contains(&"word/diagrams/layout1.xml".to_string()));
    }

    #[test]
    fn test_parse_custom_xml_part() {
        let xml = r#"<root xmlns="urn:example" xmlns:x="urn:x"><x:child/></root>"#;
        let part = parse_custom_xml_part(xml, "customXml/item1.xml", 42).expect("custom xml");
        assert_eq!(part.root_element.as_deref(), Some("root"));
        assert!(part.namespaces.iter().any(|n| n == "urn:example"));
    }

    #[test]
    fn test_build_relationship_graph() {
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="t1" Target="doc.xml"/>
          <Relationship Id="rId2" Type="t2" Target="http://example.com" TargetMode="External"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels parse");
        let graph =
            build_relationship_graph("word/document.xml", "word/_rels/document.xml.rels", &rels);
        assert_eq!(graph.relationships.len(), 2);
        assert!(graph
            .relationships
            .iter()
            .any(|r| matches!(r.target_mode, RelationshipTargetMode::External)));
    }
}
