use super::{
    ParseError, PresentationProperties, PresentationTag, ShapeType, SmartArtPart, SourceSpan,
    TableStyle, TableStyleSet, ViewProperties,
};
use crate::xml_utils::{lossy_attr_value, xml_error};
use docir_core::types::NodeId;
use quick_xml::events::Event;
use quick_xml::Reader;

pub(super) fn parse_presentation_properties(
    xml: &str,
    path: &str,
) -> Result<PresentationProperties, ParseError> {
    let mut props = PresentationProperties::new();
    props.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"p:presentationPr" {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"autoCompressPictures" => {
                                let value = lossy_attr_value(&attr);
                                props.auto_compress_pictures =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"compatMode" => {
                                props.compat_mode = Some(lossy_attr_value(&attr).to_string());
                            }
                            b"rtl" => {
                                let value = lossy_attr_value(&attr);
                                props.rtl =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"showSpecialPlsOnTitleSld" => {
                                let value = lossy_attr_value(&attr);
                                props.show_special_placeholders =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"removePersonalInfoOnSave" => {
                                let value = lossy_attr_value(&attr);
                                props.remove_personal_info_on_save =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"showInkAnnotation" => {
                                let value = lossy_attr_value(&attr);
                                props.show_ink_annotation =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(props)
}

pub(super) fn parse_view_properties(xml: &str, path: &str) -> Result<ViewProperties, ParseError> {
    let mut props = ViewProperties::new();
    props.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"p:viewPr" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"lastView" => {
                                props.last_view = Some(lossy_attr_value(&attr).to_string());
                            }
                            b"showComments" => {
                                let value = lossy_attr_value(&attr);
                                props.show_comments =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"showHiddenSlides" => {
                                let value = lossy_attr_value(&attr);
                                props.show_hidden_slides =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"showGuides" => {
                                let value = lossy_attr_value(&attr);
                                props.show_guides =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"showGrid" => {
                                let value = lossy_attr_value(&attr);
                                props.show_grid =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"showOutlineIcons" => {
                                let value = lossy_attr_value(&attr);
                                props.show_outline_icons =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            _ => {}
                        }
                    }
                }
                b"p:zoom" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"percent" {
                            props.zoom = lossy_attr_value(&attr).parse::<u32>().ok();
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(props)
}

pub(super) fn parse_table_styles(xml: &str, path: &str) -> Result<TableStyleSet, ParseError> {
    let mut styles = TableStyleSet::new();
    styles.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"a:tblStyleLst" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"def" {
                            styles.default_style_id = Some(lossy_attr_value(&attr).to_string());
                        }
                    }
                }
                b"a:tblStyle" => {
                    let mut style_id = None;
                    let mut name = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"styleId" => {
                                style_id = Some(lossy_attr_value(&attr).to_string());
                            }
                            b"name" => name = Some(lossy_attr_value(&attr).to_string()),
                            _ => {}
                        }
                    }
                    if let Some(style_id) = style_id {
                        styles.styles.push(TableStyle { style_id, name });
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(styles)
}

pub(super) fn parse_presentation_tags(
    xml: &str,
    path: &str,
) -> Result<Vec<PresentationTag>, ParseError> {
    let mut tags: Vec<PresentationTag> = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref().ends_with(b"tag") {
                    let mut name = None;
                    let mut val = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => name = Some(lossy_attr_value(&attr).to_string()),
                            b"val" => val = Some(lossy_attr_value(&attr).to_string()),
                            _ => {}
                        }
                    }
                    if let Some(name) = name {
                        tags.push(PresentationTag {
                            id: NodeId::new(),
                            name,
                            value: val,
                            span: Some(SourceSpan::new(path)),
                        });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(tags)
}

pub(super) fn parse_smartart_part(xml: &str, path: &str) -> Result<SmartArtPart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut root = None;
    let mut point_count: u32 = 0;
    let mut connection_count: u32 = 0;
    let mut rel_ids: Vec<String> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if root.is_none() {
                    root = Some(String::from_utf8_lossy(e.name().as_ref()).to_string());
                }
                let name_buf = e.name().as_ref().to_vec();
                let name = name_buf.as_slice();
                if name.ends_with(b":pt") || name == b"dgm:pt" {
                    point_count += 1;
                }
                if name.ends_with(b":cxn") || name == b"dgm:cxn" {
                    connection_count += 1;
                }
                if name.ends_with(b":relIds") || name == b"dgm:relIds" {
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        if key == b"r:dm" || key == b"r:lo" || key == b"r:qs" || key == b"r:cs" {
                            let val = lossy_attr_value(&attr).to_string();
                            if !val.is_empty() {
                                rel_ids.push(val);
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    let kind = if path.contains("layout") {
        "layout"
    } else if path.contains("style") {
        "style"
    } else if path.contains("colors") {
        "colors"
    } else {
        "data"
    };

    Ok(SmartArtPart {
        id: NodeId::new(),
        kind: kind.to_string(),
        path: path.to_string(),
        root_element: root,
        point_count: if point_count > 0 {
            Some(point_count)
        } else {
            None
        },
        connection_count: if connection_count > 0 {
            Some(connection_count)
        } else {
            None
        },
        rel_ids,
        span: Some(SourceSpan::new(path)),
    })
}

#[derive(Default)]
pub(super) struct SlideMasterMeta {
    pub(super) preserve: Option<bool>,
    pub(super) show_master_sp: Option<bool>,
    pub(super) show_master_ph_anim: Option<bool>,
}

pub(super) fn parse_slide_master_meta(
    xml: &str,
    path: &str,
) -> Result<SlideMasterMeta, ParseError> {
    let mut meta = SlideMasterMeta::default();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"p:sldMaster" {
                    for attr in e.attributes().flatten() {
                        let value = lossy_attr_value(&attr);
                        match attr.key.as_ref() {
                            b"preserve" => {
                                meta.preserve =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"showMasterSp" => {
                                meta.show_master_sp =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"showMasterPhAnim" => {
                                meta.show_master_ph_anim =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            _ => {}
                        }
                    }
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(meta)
}

#[derive(Default)]
pub(super) struct SlideLayoutMeta {
    pub(super) layout_type: Option<String>,
    pub(super) matching_name: Option<String>,
    pub(super) preserve: Option<bool>,
    pub(super) show_master_sp: Option<bool>,
    pub(super) show_master_ph_anim: Option<bool>,
}

pub(super) fn parse_slide_layout_meta(
    xml: &str,
    path: &str,
) -> Result<SlideLayoutMeta, ParseError> {
    let mut meta = SlideLayoutMeta::default();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"p:sldLayout" {
                    for attr in e.attributes().flatten() {
                        let value = lossy_attr_value(&attr);
                        match attr.key.as_ref() {
                            b"type" => meta.layout_type = Some(value.to_string()),
                            b"matchingName" => meta.matching_name = Some(value.to_string()),
                            b"preserve" => {
                                meta.preserve =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"showMasterSp" => {
                                meta.show_master_sp =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"showMasterPhAnim" => {
                                meta.show_master_ph_anim =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            _ => {}
                        }
                    }
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(meta)
}

pub(super) fn map_shape_type(value: &str) -> ShapeType {
    match value {
        "rect" => ShapeType::Rectangle,
        "roundRect" => ShapeType::RoundRect,
        "ellipse" => ShapeType::Ellipse,
        "triangle" => ShapeType::Triangle,
        "line" => ShapeType::Line,
        "arrow" => ShapeType::Arrow,
        _ => ShapeType::Custom,
    }
}
