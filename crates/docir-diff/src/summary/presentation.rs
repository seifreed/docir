use docir_core::ir::IRNode;
use docir_core::visitor::IrStore;

use super::{abbreviate, opt_bool, opt_str, opt_u32, summarize_shape, summarize_slide};

pub(crate) fn summarize(node: &IRNode, _store: &IrStore) -> Option<String> {
    match node {
        IRNode::Slide(slide) => Some(summarize_slide(slide)),
        IRNode::Shape(shape) => Some(summarize_shape(shape)),
        IRNode::SlideMaster(master) => Some(format!(
            "shapes={} layouts={}",
            master.shapes.len(),
            master.layouts.len()
        )),
        IRNode::SlideLayout(layout) => Some(format!("shapes={}", layout.shapes.len())),
        IRNode::NotesMaster(master) => Some(format!("shapes={}", master.shapes.len())),
        IRNode::HandoutMaster(master) => Some(format!("shapes={}", master.shapes.len())),
        IRNode::NotesSlide(slide) => Some(format!(
            "shapes={} text={}",
            slide.shapes.len(),
            opt_str(&slide.text)
        )),
        IRNode::PresentationProperties(props) => Some(format!(
            "auto_compress={} compat={} rtl={}",
            opt_bool(props.auto_compress_pictures),
            opt_str(&props.compat_mode),
            opt_bool(props.rtl)
        )),
        IRNode::ViewProperties(props) => Some(format!(
            "last_view={} zoom={}",
            opt_str(&props.last_view),
            opt_u32(props.zoom)
        )),
        IRNode::TableStyleSet(styles) => Some(format!(
            "default={} styles={}",
            opt_str(&styles.default_style_id),
            styles.styles.len()
        )),
        IRNode::PptxCommentAuthor(author) => Some(format!(
            "author_id={} name={}",
            author.author_id,
            opt_str(&author.name)
        )),
        IRNode::PptxComment(comment) => Some(format!(
            "author_id={} text={}",
            comment
                .author_id
                .map(|v| v.to_string())
                .unwrap_or_else(|| "-".to_string()),
            abbreviate(&comment.text, 80)
        )),
        IRNode::PresentationTag(tag) => {
            Some(format!("name={} value={}", tag.name, opt_str(&tag.value)))
        }
        IRNode::PresentationInfo(info) => Some(format!(
            "slide_size={} notes_size={} show_type={}",
            info.slide_size
                .as_ref()
                .map(|s| format!("{}x{}", s.cx, s.cy))
                .unwrap_or_else(|| "-".to_string()),
            info.notes_size
                .as_ref()
                .map(|s| format!("{}x{}", s.cx, s.cy))
                .unwrap_or_else(|| "-".to_string()),
            opt_str(&info.show_type)
        )),
        _ => None,
    }
}
