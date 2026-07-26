use super::*;
use quick_xml::Reader;

fn parse_page_start(xml: &[u8]) -> (Reader<std::io::Cursor<&[u8]>>, BytesStart<'static>) {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let page_start = loop {
        match reader.read_event_into(&mut buf).unwrap() {
            Event::Start(e) if local_name(e.name().as_ref()) == b"page" => {
                break e.into_owned();
            }
            Event::Eof => panic!("missing draw:page"),
            _ => {}
        }
        buf.clear();
    };
    (reader, page_start)
}

#[test]
fn parse_draw_page_extracts_metadata_transition_notes_and_shape_text() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"
  xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"
  xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <d:page d:name="SlideA"
    p:master-page-name="MasterA"
    d:style-name="LayoutFallback"
    p:transition-type="fade"
    p:transition-duration="2"
    p:animation="click">
    <d:frame d:name="TitleShape">
      <d:text-box>
        <t:p>Hello ODP</t:p>
      </d:text-box>
    </d:frame>
    <p:notes>
      <t:p>Speaker notes</t:p>
    </p:notes>
    <anim:animate smil:targetElement="TitleShape" smil:dur="PT1S"/>
  </d:page>
</office:document-content>
"#;

    let (mut reader, page_start) = parse_page_start(xml);
    let mut store = IrStore::new();

    let slide = parse_draw_page(&mut reader, &page_start, 3, &mut store).unwrap();
    assert_eq!(slide.number, 3);
    assert_eq!(slide.name.as_deref(), Some("SlideA"));
    assert_eq!(slide.master_id.as_deref(), Some("MasterA"));
    assert_eq!(slide.layout_id.as_deref(), Some("LayoutFallback"));
    assert_eq!(slide.notes.as_deref(), Some("Speaker notes"));
    assert_eq!(slide.shapes.len(), 1);

    let transition = slide.transition.expect("expected transition");
    assert_eq!(transition.transition_type.as_deref(), Some("fade"));
    assert_eq!(transition.duration_ms, Some(2));
    assert_eq!(transition.advance_on_click, Some(true));
    assert_eq!(slide.animations.len(), 1);
    assert_eq!(slide.animations[0].target.as_deref(), Some("TitleShape"));
    assert_eq!(slide.animations[0].duration_ms, Some(1000));

    let Some(IRNode::Shape(shape)) = store.get(slide.shapes[0]) else {
        panic!("expected shape node");
    };
    assert_eq!(shape.name.as_deref(), Some("TitleShape"));
    assert_eq!(shape.shape_type, ShapeType::TextBox);
    let text = shape.text.as_ref().expect("expected shape text");
    assert_eq!(text.paragraphs.len(), 1);
    assert_eq!(text.paragraphs[0].runs[0].text, "Hello ODP");
}

#[test]
fn parse_draw_frame_presentation_returns_none_for_unrecognized_content() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<draw:frame xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">
  <draw:unknown/>
</draw:frame>
"#;
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let frame_start = loop {
        match reader.read_event_into(&mut buf).unwrap() {
            Event::Start(e) if e.name().as_ref() == b"draw:frame" => break e.into_owned(),
            Event::Eof => panic!("missing draw:frame"),
            _ => {}
        }
        buf.clear();
    };
    let mut store = IrStore::new();

    let shape = parse_draw_frame_presentation(&mut reader, &frame_start, &mut store).unwrap();
    assert!(shape.is_none());
    assert_eq!(store.values().count(), 0);
}

#[test]
fn parse_draw_frame_presentation_classifies_plugin_media() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<draw:frame xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  draw:name="ClipFrame">
  <draw:plugin xlink:href="media/clip.mp4"/>
</draw:frame>
"#;
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let frame_start = loop {
        match reader.read_event_into(&mut buf).unwrap() {
            Event::Start(e) if e.name().as_ref() == b"draw:frame" => break e.into_owned(),
            Event::Eof => panic!("missing draw:frame"),
            _ => {}
        }
        buf.clear();
    };
    let mut store = IrStore::new();

    let shape_id = parse_draw_frame_presentation(&mut reader, &frame_start, &mut store).unwrap();
    let Some(shape_id) = shape_id else {
        panic!("expected shape");
    };
    let Some(IRNode::Shape(shape)) = store.get(shape_id) else {
        panic!("expected shape node");
    };
    assert_eq!(shape.name.as_deref(), Some("ClipFrame"));
    assert_eq!(shape.media_target.as_deref(), Some("media/clip.mp4"));
    assert_eq!(shape.shape_type, ShapeType::Video);
}

#[test]
fn parse_draw_frame_presentation_rejects_malformed_name_attributes() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<draw:frame xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  draw:name="ClipFrame" draw:name="Duplicate">
  <draw:plugin xlink:href="media/clip.mp4"/>
</draw:frame>
"#;
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let frame_start = loop {
        match reader.read_event_into(&mut buf).unwrap() {
            Event::Start(e) if e.name().as_ref() == b"draw:frame" => break e.into_owned(),
            Event::Eof => panic!("missing draw:frame"),
            _ => {}
        }
        buf.clear();
    };
    let mut store = IrStore::new();

    let err = parse_draw_frame_presentation(&mut reader, &frame_start, &mut store)
        .expect_err("duplicate frame name attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn parse_custom_shape_presentation_preserves_text_runs() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<draw:custom-shape xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  draw:name="Badge">
  <text:p>First line</text:p>
  <text:p>Second line</text:p>
</draw:custom-shape>
"#;
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let shape_start = loop {
        match reader.read_event_into(&mut buf).unwrap() {
            Event::Start(e) if e.name().as_ref() == b"draw:custom-shape" => {
                break e.into_owned();
            }
            Event::Eof => panic!("missing draw:custom-shape"),
            _ => {}
        }
        buf.clear();
    };
    let mut store = IrStore::new();

    let shape_id = parse_custom_shape_presentation(&mut reader, &shape_start, &mut store).unwrap();
    let Some(shape_id) = shape_id else {
        panic!("expected custom shape");
    };
    let Some(IRNode::Shape(shape)) = store.get(shape_id) else {
        panic!("expected shape node");
    };
    assert_eq!(shape.name.as_deref(), Some("Badge"));
    assert_eq!(shape.shape_type, ShapeType::Custom);
    let text = shape.text.as_ref().expect("shape text");
    assert_eq!(text.paragraphs.len(), 2);
    assert_eq!(text.paragraphs[0].runs[0].text, "First line");
    assert_eq!(text.paragraphs[1].runs[0].text, "Second line");
}

#[test]
fn parse_custom_shape_presentation_rejects_malformed_name_attributes() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<draw:custom-shape xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  draw:name="Badge" draw:name="Duplicate">
  <text:p>First line</text:p>
</draw:custom-shape>
"#;
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let shape_start = loop {
        match reader.read_event_into(&mut buf).unwrap() {
            Event::Start(e) if e.name().as_ref() == b"draw:custom-shape" => {
                break e.into_owned();
            }
            Event::Eof => panic!("missing draw:custom-shape"),
            _ => {}
        }
        buf.clear();
    };
    let mut store = IrStore::new();

    let err = parse_custom_shape_presentation(&mut reader, &shape_start, &mut store)
        .expect_err("duplicate custom shape name attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn classify_media_shape_covers_audio_video_and_unknown_paths() {
    assert_eq!(classify_media_shape("media/clip.ogg"), ShapeType::Audio);
    assert_eq!(classify_media_shape("media/clip.OGV"), ShapeType::Video);
    assert_eq!(classify_media_shape("media/blob.bin"), ShapeType::Unknown);
}

#[test]
fn parse_odf_animation_prefers_target_fallback_and_parses_iso_duration() {
    let mut start = BytesStart::new("anim:animate");
    start.push_attribute(("smil:targetElement", "shape-42"));
    start.push_attribute(("smil:dur", "PT2.5S"));
    start.push_attribute(("presentation:preset-id", "entrance"));

    let anim = parse_odf_animation(&start)
        .expect("animation parse")
        .expect("animation metadata");
    assert_eq!(anim.target.as_deref(), Some("shape-42"));
    assert_eq!(anim.duration_ms, Some(2500));
    assert_eq!(anim.preset_id.as_deref(), Some("entrance"));
}

#[test]
fn parse_draw_page_rejects_malformed_animation_attributes() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"
  xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0">
  <d:page d:name="SlideA">
    <anim:animate smil:targetElement="shape-1" smil:targetElement="shape-2"/>
  </d:page>
</office:document-content>
"#;

    let (mut reader, page_start) = parse_page_start(xml);
    let mut store = IrStore::new();

    let err = parse_draw_page(&mut reader, &page_start, 1, &mut store)
        .expect_err("duplicate animation attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn parse_draw_page_accepts_alternate_animation_prefix() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:a="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"
  xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0">
  <d:page d:name="SlideA">
    <a:animate smil:targetElement="shape-1" smil:dur="PT1S"/>
  </d:page>
</office:document-content>
"#;

    let (mut reader, page_start) = parse_page_start(xml);
    let mut store = IrStore::new();

    let slide = parse_draw_page(&mut reader, &page_start, 1, &mut store).expect("slide");

    assert_eq!(slide.animations.len(), 1);
    assert_eq!(slide.animations[0].target.as_deref(), Some("shape-1"));
    assert_eq!(slide.animations[0].duration_ms, Some(1000));
}

#[test]
fn parse_odp_transition_supports_fallback_keys_and_ignores_advance_only() {
    let mut fallback = BytesStart::new("draw:page");
    fallback.push_attribute(("draw:transition-type", "wipe"));
    fallback.push_attribute(("presentation:transition-speed", "fast"));
    fallback.push_attribute(("presentation:transition-duration", "900"));
    fallback.push_attribute(("presentation:duration", "1200"));
    fallback.push_attribute(("presentation:animation", "click"));

    let transition = parse_odp_transition(&fallback)
        .expect("transition parse")
        .expect("transition");
    assert_eq!(transition.transition_type.as_deref(), Some("wipe"));
    assert_eq!(transition.speed.as_deref(), Some("fast"));
    assert_eq!(transition.duration_ms, Some(900));
    assert_eq!(transition.advance_after_ms, Some(1200));
    assert_eq!(transition.advance_on_click, Some(true));

    let mut advance_only = BytesStart::new("draw:page");
    advance_only.push_attribute(("presentation:duration", "1000"));
    assert!(
        parse_odp_transition(&advance_only)
            .expect("transition parse")
            .is_none()
    );
}

#[test]
fn parse_odp_transition_reports_malformed_numeric_attributes() {
    for attr in ["presentation:transition-duration", "presentation:duration"] {
        let mut start = BytesStart::new("draw:page");
        start.push_attribute(("presentation:transition-type", "fade"));
        start.push_attribute((attr, "bad"));

        let err = parse_odp_transition(&start).expect_err("malformed transition number must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }
}

#[test]
fn parse_duration_and_media_classification_cover_helper_paths() {
    assert_eq!(parse_duration_ms("250ms"), Some(250));
    assert_eq!(parse_duration_ms("1.25s"), Some(1250));
    assert_eq!(parse_duration_ms("PT3S"), Some(3000));
    assert_eq!(parse_duration_ms("PTPT3S"), None);
    assert_eq!(parse_duration_ms("PT3SS"), None);
    assert_eq!(parse_duration_ms("NaNs"), None);
    assert_eq!(parse_duration_ms("PTinfS"), None);
    assert_eq!(parse_duration_ms("-1s"), None);
    assert_eq!(parse_duration_ms(""), None);
    assert_eq!(parse_duration_ms("invalid"), None);

    assert_eq!(
        classify_media_type("media/clip.oga", "application/ogg"),
        Some(MediaType::Audio)
    );
    assert_eq!(
        classify_media_type("media/clip.ogv", "application/ogg"),
        Some(MediaType::Video)
    );
    assert_eq!(
        classify_media_type("media/blob.bin", "application/octet-stream"),
        None
    );

    let asset = build_media_asset("Pictures/p1.png", "image/png", 42).expect("asset");
    assert_eq!(asset.media_type, MediaType::Image);
    assert_eq!(asset.content_type.as_deref(), Some("image/png"));
    assert_eq!(
        asset.span.as_ref().map(|s| s.file_path.as_str()),
        Some("META-INF/manifest.xml")
    );
}

#[test]
fn parse_frame_shape_empty_covers_image_chart_and_ole_variants() {
    let mut frame = FrameShapeState::new();
    let mut store = IrStore::new();

    let mut image = BytesStart::new("dr:image");
    image.push_attribute(("lnk:href", "Pictures/img1.png"));
    parse_frame_shape_empty(&image, &mut store, &mut frame).expect("image frame shape");
    assert_eq!(frame.shape_type, ShapeType::Picture);
    assert_eq!(frame.media_target.as_deref(), Some("Pictures/img1.png"));
    assert!(frame.has_shape);

    let mut chart = BytesStart::new("ch:chart");
    chart.push_attribute(("ch:class", "bar"));
    parse_frame_shape_empty(&chart, &mut store, &mut frame).expect("chart frame shape");
    assert_eq!(frame.shape_type, ShapeType::Chart);
    let chart_id = frame.chart_id.expect("chart id");
    let Some(IRNode::ChartData(chart_node)) = store.get(chart_id) else {
        panic!("expected chart data node");
    };
    assert_eq!(chart_node.chart_type.as_deref(), Some("bar"));

    let object = BytesStart::new("dr:object-ole");
    parse_frame_shape_empty(&object, &mut store, &mut frame).expect("ole frame shape");
    assert_eq!(frame.shape_type, ShapeType::OleObject);
    assert!(frame.has_shape);
}

#[test]
fn parse_frame_shape_start_extracts_chart_title_text() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<chart:chart xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  chart:class="chart:line">
  <chart:title>
    <text:p>Main</text:p>
    <text:p>Title</text:p>
  </chart:title>
</chart:chart>
"#;
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let start = loop {
        match reader.read_event_into(&mut buf).unwrap() {
            Event::Start(e) if e.name().as_ref() == b"chart:chart" => break e.into_owned(),
            Event::Eof => panic!("missing chart:chart"),
            _ => {}
        }
        buf.clear();
    };

    let mut frame = FrameShapeState::new();
    let mut store = IrStore::new();
    parse_frame_shape_start(&mut reader, &start, &mut store, &mut frame).expect("parse");

    assert_eq!(frame.shape_type, ShapeType::Chart);
    assert!(frame.has_shape);
    let chart_id = frame.chart_id.expect("chart id");
    let Some(IRNode::ChartData(chart)) = store.get(chart_id) else {
        panic!("expected chart data");
    };
    assert_eq!(chart.chart_type.as_deref(), Some("chart:line"));
    assert_eq!(chart.title.as_deref(), Some("Main Title"));
}

#[test]
fn parse_frame_shape_start_rejects_malformed_chart_class_attributes() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<chart:chart xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
  chart:class="chart:line" chart:class="chart:bar">
</chart:chart>
"#;
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let start = loop {
        match reader.read_event_into(&mut buf).unwrap() {
            Event::Start(e) if e.name().as_ref() == b"chart:chart" => break e.into_owned(),
            Event::Eof => panic!("missing chart:chart"),
            _ => {}
        }
        buf.clear();
    };

    let mut frame = FrameShapeState::new();
    let mut store = IrStore::new();
    let err = parse_frame_shape_start(&mut reader, &start, &mut store, &mut frame)
        .expect_err("duplicate chart class attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}
