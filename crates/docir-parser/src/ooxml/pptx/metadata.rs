use super::{
    ParseError, PresentationProperties, PresentationTag, ShapeType, SmartArtPart, SourceSpan,
    TableStyle, TableStyleSet, ViewProperties,
};
use crate::xml_utils::parse_bool_attr;
use crate::xml_utils::{local_name, lossy_attr_value, track_xml_document_event, xml_error};
use docir_core::types::NodeId;
use quick_xml::Reader;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::events::attributes::Attribute;

fn visit_attributes<F>(start: &BytesStart<'_>, path: &str, mut visit: F) -> Result<(), ParseError>
where
    F: FnMut(&Attribute<'_>) -> Result<(), ParseError>,
{
    for attr in start.attributes() {
        let attr = attr.map_err(|err| xml_error(path, err))?;
        visit(&attr)?;
    }
    Ok(())
}

pub(super) fn parse_presentation_properties(
    xml: &str,
    path: &str,
) -> Result<PresentationProperties, ParseError> {
    let mut props = PresentationProperties::new();
    props.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut root_closed = false;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error(path, err))?;
        if track_xml_document_event(&event, &mut depth, &mut root_closed, path)? {
            break;
        }
        match event {
            Event::Start(e) | Event::Empty(e)
                if local_name(e.name().as_ref()) == b"presentationPr" =>
            {
                visit_attributes(&e, path, |attr| match attr.key.as_ref() {
                    b"autoCompressPictures" => {
                        let value = lossy_attr_value(attr);
                        props.auto_compress_pictures =
                            Some(parse_bool_attr(value.as_bytes(), path)?);
                        Ok(())
                    }
                    b"compatMode" => {
                        props.compat_mode = Some(lossy_attr_value(attr).to_string());
                        Ok(())
                    }
                    b"rtl" => {
                        let value = lossy_attr_value(attr);
                        props.rtl = Some(parse_bool_attr(value.as_bytes(), path)?);
                        Ok(())
                    }
                    b"showSpecialPlsOnTitleSld" => {
                        let value = lossy_attr_value(attr);
                        props.show_special_placeholders =
                            Some(parse_bool_attr(value.as_bytes(), path)?);
                        Ok(())
                    }
                    b"removePersonalInfoOnSave" => {
                        let value = lossy_attr_value(attr);
                        props.remove_personal_info_on_save =
                            Some(parse_bool_attr(value.as_bytes(), path)?);
                        Ok(())
                    }
                    b"showInkAnnotation" => {
                        let value = lossy_attr_value(attr);
                        props.show_ink_annotation = Some(parse_bool_attr(value.as_bytes(), path)?);
                        Ok(())
                    }
                    _ => Ok(()),
                })?;
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
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut root_closed = false;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error(path, err))?;
        if track_xml_document_event(&event, &mut depth, &mut root_closed, path)? {
            break;
        }
        match event {
            Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"viewPr" => {
                    visit_attributes(&e, path, |attr| match attr.key.as_ref() {
                        b"lastView" => {
                            props.last_view = Some(lossy_attr_value(attr).to_string());
                            Ok(())
                        }
                        b"showComments" => {
                            let value = lossy_attr_value(attr);
                            props.show_comments = Some(parse_bool_attr(value.as_bytes(), path)?);
                            Ok(())
                        }
                        b"showHiddenSlides" => {
                            let value = lossy_attr_value(attr);
                            props.show_hidden_slides =
                                Some(parse_bool_attr(value.as_bytes(), path)?);
                            Ok(())
                        }
                        b"showGuides" => {
                            let value = lossy_attr_value(attr);
                            props.show_guides = Some(parse_bool_attr(value.as_bytes(), path)?);
                            Ok(())
                        }
                        b"showGrid" => {
                            let value = lossy_attr_value(attr);
                            props.show_grid = Some(parse_bool_attr(value.as_bytes(), path)?);
                            Ok(())
                        }
                        b"showOutlineIcons" => {
                            let value = lossy_attr_value(attr);
                            props.show_outline_icons =
                                Some(parse_bool_attr(value.as_bytes(), path)?);
                            Ok(())
                        }
                        _ => Ok(()),
                    })?;
                }
                b"zoom" => {
                    visit_attributes(&e, path, |attr| {
                        if attr.key.as_ref() == b"percent" {
                            props.zoom = Some(
                                lossy_attr_value(attr)
                                    .parse::<u32>()
                                    .map_err(|err| xml_error(path, err))?,
                            );
                        }
                        Ok(())
                    })?;
                }
                _ => {}
            },
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
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut root_closed = false;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error(path, err))?;
        if track_xml_document_event(&event, &mut depth, &mut root_closed, path)? {
            break;
        }
        match event {
            Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"tblStyleLst" => {
                    visit_attributes(&e, path, |attr| {
                        if attr.key.as_ref() == b"def" {
                            styles.default_style_id = Some(lossy_attr_value(attr).to_string());
                        }
                        Ok(())
                    })?;
                }
                b"tblStyle" => {
                    let mut style_id = None;
                    let mut name = None;
                    visit_attributes(&e, path, |attr| match attr.key.as_ref() {
                        b"styleId" => {
                            style_id = Some(lossy_attr_value(attr).to_string());
                            Ok(())
                        }
                        b"name" => {
                            name = Some(lossy_attr_value(attr).to_string());
                            Ok(())
                        }
                        _ => Ok(()),
                    })?;
                    if let Some(style_id) = style_id {
                        styles.styles.push(TableStyle { style_id, name });
                    }
                }
                _ => {}
            },
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
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut root_closed = false;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error(path, err))?;
        if track_xml_document_event(&event, &mut depth, &mut root_closed, path)? {
            break;
        }
        match event {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"tag" => {
                let mut name = None;
                let mut val = None;
                visit_attributes(&e, path, |attr| match attr.key.as_ref() {
                    b"name" => {
                        name = Some(lossy_attr_value(attr).to_string());
                        Ok(())
                    }
                    b"val" => {
                        val = Some(lossy_attr_value(attr).to_string());
                        Ok(())
                    }
                    _ => Ok(()),
                })?;
                if let Some(name) = name {
                    tags.push(PresentationTag {
                        id: NodeId::new(),
                        name,
                        value: val,
                        span: Some(SourceSpan::new(path)),
                    });
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(tags)
}

pub(super) fn parse_smartart_part(xml: &str, path: &str) -> Result<SmartArtPart, ParseError> {
    let mut reader = Reader::from_str(xml);
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut root = None;
    let mut point_count: u32 = 0;
    let mut connection_count: u32 = 0;
    let mut rel_ids: Vec<String> = Vec::new();
    let mut depth = 0usize;
    let mut root_closed = false;
    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error(path, err))?;
        if track_xml_document_event(&event, &mut depth, &mut root_closed, path)? {
            break;
        }
        match event {
            Event::Start(e) | Event::Empty(e) => {
                if root.is_none() {
                    root = Some(String::from_utf8_lossy(e.name().as_ref()).to_string());
                }
                let name_buf = e.name().as_ref().to_vec();
                let name = local_name(name_buf.as_slice());
                if name == b"pt" {
                    point_count += 1;
                }
                if name == b"cxn" {
                    connection_count += 1;
                }
                if name == b"relIds" {
                    visit_attributes(&e, path, |attr| {
                        let key = local_name(attr.key.as_ref());
                        if matches!(key, b"dm" | b"lo" | b"qs" | b"cs") {
                            let val = lossy_attr_value(attr).to_string();
                            if !val.is_empty() {
                                rel_ids.push(val);
                            }
                        }
                        Ok(())
                    })?;
                }
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
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut root_closed = false;
    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error(path, err))?;
        if track_xml_document_event(&event, &mut depth, &mut root_closed, path)? {
            break;
        }
        match event {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"sldMaster" => {
                visit_attributes(&e, path, |attr| {
                    let value = lossy_attr_value(attr);
                    match attr.key.as_ref() {
                        b"preserve" => {
                            meta.preserve = Some(parse_bool_attr(value.as_bytes(), path)?);
                            Ok(())
                        }
                        b"showMasterSp" => {
                            meta.show_master_sp = Some(parse_bool_attr(value.as_bytes(), path)?);
                            Ok(())
                        }
                        b"showMasterPhAnim" => {
                            meta.show_master_ph_anim =
                                Some(parse_bool_attr(value.as_bytes(), path)?);
                            Ok(())
                        }
                        _ => Ok(()),
                    }
                })?;
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
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut root_closed = false;
    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error(path, err))?;
        if track_xml_document_event(&event, &mut depth, &mut root_closed, path)? {
            break;
        }
        match event {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"sldLayout" => {
                visit_attributes(&e, path, |attr| {
                    let value = lossy_attr_value(attr);
                    match attr.key.as_ref() {
                        b"type" => {
                            meta.layout_type = Some(value.to_string());
                            Ok(())
                        }
                        b"matchingName" => {
                            meta.matching_name = Some(value.to_string());
                            Ok(())
                        }
                        b"preserve" => {
                            meta.preserve = Some(parse_bool_attr(value.as_bytes(), path)?);
                            Ok(())
                        }
                        b"showMasterSp" => {
                            meta.show_master_sp = Some(parse_bool_attr(value.as_bytes(), path)?);
                            Ok(())
                        }
                        b"showMasterPhAnim" => {
                            meta.show_master_ph_anim =
                                Some(parse_bool_attr(value.as_bytes(), path)?);
                            Ok(())
                        }
                        _ => Ok(()),
                    }
                })?;
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
