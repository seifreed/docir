use super::super::*;
use super::helpers::create_pptx_with_media;

#[test]
fn test_pptx_animation_media_asset_link() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld><p:spTree/></p:cSld>
          <p:timing>
            <p:audio r:link="rIdAudio" dur="5000"/>
          </p:timing>
        </p:sld>
        "#;
    let slide_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdAudio"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
            Target="../media/audio1.wav"/>
        </Relationships>
        "#;

    let path = create_pptx_with_media(slide_xml, slide_rels);
    let parser = OoxmlParser::new();
    let parsed = parser.parse_file(&path).expect("parse pptx");
    let doc = parsed.document().expect("doc");
    let slide_id = doc.content[0];
    let slide = match parsed.store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    assert_eq!(slide.animations.len(), 1);
    let anim = &slide.animations[0];
    let media_id = anim.media_asset.expect("media asset link");
    let media = match parsed.store.get(media_id) {
        Some(IRNode::MediaAsset(m)) => m,
        _ => panic!("missing media asset"),
    };
    assert_eq!(media.path, "ppt/media/audio1.wav");
    std::fs::remove_file(path).ok();
}
