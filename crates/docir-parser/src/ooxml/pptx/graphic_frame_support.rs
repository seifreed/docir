use super::graphic_frame::GraphicFrameState;
use crate::error::ParseError;
use crate::xml_utils::{local_name, lossy_attr_value, xml_error};
use docir_core::ir::{Shape, ShapeType};
use quick_xml::events::BytesStart;

pub(super) fn apply_non_visual_shape_props(
    shape: &mut Shape,
    e: &BytesStart<'_>,
    slide_path: &str,
) -> Result<(), ParseError> {
    for attr in e.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
        match attr.key.as_ref() {
            b"name" => {
                shape.name = Some(lossy_attr_value(&attr).to_string());
            }
            b"descr" => {
                shape.alt_text = Some(lossy_attr_value(&attr).to_string());
            }
            _ => {}
        }
    }
    Ok(())
}

pub(super) fn apply_graphic_data_shape_type(
    shape: &mut Shape,
    e: &BytesStart<'_>,
    slide_path: &str,
) -> Result<(), ParseError> {
    for attr in e.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
        if attr.key.as_ref() == b"uri" {
            let uri = lossy_attr_value(&attr);
            if uri.contains("chart") {
                shape.shape_type = ShapeType::Chart;
            } else if uri.contains("table") {
                shape.shape_type = ShapeType::Table;
            } else if uri.contains("ole") || uri.contains("object") {
                shape.shape_type = ShapeType::OleObject;
            }
        }
    }
    Ok(())
}

pub(super) fn capture_chart_rel(
    state: &mut GraphicFrameState,
    e: &BytesStart<'_>,
    slide_path: &str,
) -> Result<(), ParseError> {
    for attr in e.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
        if state.chart_rel.is_none() && local_name(attr.key.as_ref()) == b"id" {
            let val = lossy_attr_value(&attr).to_string();
            if val.starts_with("rId") {
                state.chart_rel = Some(val);
            }
        }
    }
    Ok(())
}

pub(super) fn capture_ole_rel(
    state: &mut GraphicFrameState,
    e: &BytesStart<'_>,
    slide_path: &str,
) -> Result<(), ParseError> {
    for attr in e.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
        if local_name(attr.key.as_ref()) == b"id" {
            state.ole_rel = Some(lossy_attr_value(&attr).to_string());
        }
    }
    Ok(())
}
