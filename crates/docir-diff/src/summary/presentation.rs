use docir_core::ir::IRNode;
use docir_core::visitor::IrStore;

use super::summary_parse::{
    abbreviate, opt_bool, opt_str, opt_u32, summarize_shape, summarize_slide,
};

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

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::{
        HandoutMaster, NotesMaster, NotesSlide, PptxComment, PptxCommentAuthor, PresentationInfo,
        PresentationProperties, PresentationTag, Shape, ShapeType, Slide, SlideLayout, SlideMaster,
        SlideSize, TableStyleSet, ViewProperties,
    };
    use docir_core::types::NodeId;
    use docir_core::visitor::IrStore;

    #[test]
    fn summarizes_slide_and_shape_nodes() {
        let store = IrStore::new();

        let mut slide = Slide::new(3);
        slide.name = Some("Agenda".to_string());
        slide.hidden = true;
        assert!(summarize(&IRNode::Slide(slide), &store)
            .unwrap()
            .contains("number=3"));

        let mut shape = Shape::new(ShapeType::TextBox);
        shape.name = Some("Title".to_string());
        shape.hyperlink = Some("https://example.test".to_string());
        assert!(summarize(&IRNode::Shape(shape), &store)
            .unwrap()
            .contains("name=Title"));
    }

    #[test]
    fn summarizes_layout_and_master_nodes() {
        let store = IrStore::new();
        let mut master = SlideMaster::new();
        master.shapes.push(NodeId::new());
        master.layouts.push(NodeId::new());
        assert_eq!(
            summarize(&IRNode::SlideMaster(master), &store).unwrap(),
            "shapes=1 layouts=1"
        );

        let mut layout = SlideLayout::new();
        layout.shapes.push(NodeId::new());
        assert_eq!(
            summarize(&IRNode::SlideLayout(layout), &store).unwrap(),
            "shapes=1"
        );

        let mut notes_master = NotesMaster::new();
        notes_master.shapes.push(NodeId::new());
        assert_eq!(
            summarize(&IRNode::NotesMaster(notes_master), &store).unwrap(),
            "shapes=1"
        );

        let mut handout = HandoutMaster::new();
        handout.shapes.push(NodeId::new());
        assert_eq!(
            summarize(&IRNode::HandoutMaster(handout), &store).unwrap(),
            "shapes=1"
        );
    }

    #[test]
    fn summarizes_notes_and_presentation_properties() {
        let store = IrStore::new();
        let mut notes = NotesSlide::new();
        notes.shapes.push(NodeId::new());
        notes.text = Some("Speaker notes".to_string());
        assert_eq!(
            summarize(&IRNode::NotesSlide(notes), &store).unwrap(),
            "shapes=1 text=Speaker notes"
        );

        let mut props = PresentationProperties::new();
        props.auto_compress_pictures = Some(true);
        props.compat_mode = Some("strict".to_string());
        props.rtl = Some(false);
        assert_eq!(
            summarize(&IRNode::PresentationProperties(props), &store).unwrap(),
            "auto_compress=true compat=strict rtl=false"
        );

        let mut view = ViewProperties::new();
        view.last_view = Some("sldView".to_string());
        view.zoom = Some(120);
        assert_eq!(
            summarize(&IRNode::ViewProperties(view), &store).unwrap(),
            "last_view=sldView zoom=120"
        );
    }

    #[test]
    fn summarizes_style_comment_and_tag_nodes() {
        let store = IrStore::new();
        let mut table_styles = TableStyleSet::new();
        table_styles.default_style_id = Some("TableStyleMedium2".to_string());
        table_styles.styles.push(docir_core::ir::TableStyle {
            style_id: "TableStyleMedium2".to_string(),
            name: Some("Medium".to_string()),
        });
        assert_eq!(
            summarize(&IRNode::TableStyleSet(table_styles), &store).unwrap(),
            "default=TableStyleMedium2 styles=1"
        );

        let author = PptxCommentAuthor {
            id: NodeId::new(),
            author_id: 9,
            name: Some("Reviewer".to_string()),
            initials: Some("RV".to_string()),
            span: None,
        };
        assert_eq!(
            summarize(&IRNode::PptxCommentAuthor(author), &store).unwrap(),
            "author_id=9 name=Reviewer"
        );

        let comment = PptxComment {
            id: NodeId::new(),
            author_id: Some(9),
            author_name: Some("Reviewer".to_string()),
            author_initials: Some("RV".to_string()),
            datetime: None,
            text: "Looks good".to_string(),
            span: None,
        };
        assert_eq!(
            summarize(&IRNode::PptxComment(comment), &store).unwrap(),
            "author_id=9 text=Looks good"
        );

        let tag = PresentationTag {
            id: NodeId::new(),
            name: "Env".to_string(),
            value: Some("Prod".to_string()),
            span: None,
        };
        assert_eq!(
            summarize(&IRNode::PresentationTag(tag), &store).unwrap(),
            "name=Env value=Prod"
        );
    }

    #[test]
    fn summarizes_presentation_info_and_unsupported_nodes() {
        let store = IrStore::new();
        let mut info = PresentationInfo::new();
        info.slide_size = Some(SlideSize {
            cx: 9144000,
            cy: 6858000,
            size_type: Some("screen4x3".to_string()),
        });
        info.notes_size = Some(SlideSize {
            cx: 6858000,
            cy: 9144000,
            size_type: None,
        });
        info.show_type = Some("window".to_string());
        assert_eq!(
            summarize(&IRNode::PresentationInfo(info), &store).unwrap(),
            "slide_size=9144000x6858000 notes_size=6858000x9144000 show_type=window"
        );

        assert!(summarize(
            &IRNode::Document(docir_core::ir::Document::new(
                docir_core::types::DocumentFormat::WordProcessing
            )),
            &store
        )
        .is_none());
    }
}
