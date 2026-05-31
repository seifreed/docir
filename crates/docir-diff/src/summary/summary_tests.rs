mod tests {
    use crate::summary::{IRNode, style_signature, summarize, summarize_primary};
    use docir_core::ir::{
        BookmarkEnd, BookmarkStart, Comment, ContentControl, CustomXmlPart, DiagnosticEntry,
        DiagnosticSeverity, Diagnostics, DigitalSignature, DocumentMetadata, DrawingPart, Endnote,
        ExtensionPart, ExtensionPartKind, Field, Footer, GlossaryDocument, GlossaryEntry, Header,
        Hyperlink, NumberingSet, Paragraph, PeoplePart, RelationshipGraph, Revision, RevisionType,
        Run, SmartArtPart, StyleSet, Table, TableCell, TableRow, Theme, ThemeColor, VmlDrawing,
        VmlShape, WebExtension, WebExtensionTaskpane,
    };
    use docir_core::ir::{CommentExtensionSet, CommentIdMap, CommentRangeEnd, CommentRangeStart};
    use docir_core::ir::{CommentReference, MediaAsset, MediaType};
    use docir_core::security::ExternalRefType;
    use docir_core::types::{DocumentFormat, NodeId};
    use docir_core::visitor::IrStore;

    #[test]
    fn summarizes_primary_word_and_security_nodes() {
        let store = IrStore::new();

        let doc = IRNode::Document(docir_core::ir::Document::new(
            DocumentFormat::WordProcessing,
        ));
        assert!(
            summarize_primary(&doc, &store)
                .unwrap()
                .contains("format=WordProcessing")
        );

        let mut section = docir_core::ir::Section::new();
        section.name = Some("Body".to_string());
        assert!(
            summarize_primary(&IRNode::Section(section), &store)
                .unwrap()
                .contains("name=Body")
        );

        let para_id = NodeId::new();
        let mut with_store = IrStore::new();
        let run = Run::new("Hello");
        with_store.insert(IRNode::Run(run.clone()));
        let mut para = Paragraph::new();
        para.id = para_id;
        para.runs.push(run.id);
        assert!(
            summarize_primary(&IRNode::Paragraph(para), &with_store)
                .unwrap()
                .contains("text=\"Hello\"")
        );

        let run = Run::new("Run text");
        assert!(
            summarize_primary(&IRNode::Run(run.clone()), &store)
                .unwrap()
                .contains("bold=-")
        );

        let mut link = Hyperlink::new("https://example.test", true);
        link.runs.push(run.id);
        assert!(
            summarize_primary(&IRNode::Hyperlink(link), &with_store)
                .unwrap()
                .contains("external=true")
        );

        let mut macro_project = docir_core::security::MacroProject::new();
        macro_project.name = Some("VBA".to_string());
        macro_project.has_auto_exec = true;
        assert!(
            summarize_primary(&IRNode::MacroProject(macro_project), &store)
                .unwrap()
                .contains("auto_exec=true")
        );

        let mut ext = docir_core::security::ExternalReference::new(
            ExternalRefType::Hyperlink,
            "http://example.test",
        );
        ext.ref_type = ExternalRefType::Hyperlink;
        assert!(
            summarize_primary(&IRNode::ExternalReference(ext), &store)
                .unwrap()
                .contains("target=http://example.test")
        );
    }

    #[test]
    fn summarize_covers_primary_word_and_package_nodes() {
        let store = IrStore::new();

        let mut table = Table::new();
        table.properties.style_id = Some("Grid".to_string());
        assert!(summarize(&IRNode::Table(table), &store).contains("style=Grid"));

        let mut row = TableRow::new();
        row.cells.push(NodeId::new());
        assert_eq!(summarize(&IRNode::TableRow(row), &store), "cells=1");

        let mut cell = TableCell::new();
        cell.properties.grid_span = Some(3);
        assert_eq!(
            summarize(&IRNode::TableCell(cell), &store),
            "content_nodes=0 span=3"
        );

        let mut comment = Comment::new("c-1");
        comment.author = Some("alice".to_string());
        assert!(summarize(&IRNode::Comment(comment), &store).contains("author=alice"));

        let mut footnote = docir_core::ir::Footnote::new("f-1");
        footnote.content.push(NodeId::new());
        assert_eq!(
            summarize(&IRNode::Footnote(footnote), &store),
            "id=f-1 content_nodes=1"
        );

        let mut endnote = Endnote::new("e-1");
        endnote.content.push(NodeId::new());
        assert_eq!(
            summarize(&IRNode::Endnote(endnote), &store),
            "id=e-1 content_nodes=1"
        );

        let mut header = Header::new();
        header.content.push(NodeId::new());
        assert_eq!(
            summarize(&IRNode::Header(header), &store),
            "content_nodes=1"
        );

        let mut footer = Footer::new();
        footer.content.push(NodeId::new());
        assert_eq!(
            summarize(&IRNode::Footer(footer), &store),
            "content_nodes=1"
        );

        let mut control = ContentControl::new();
        control.tag = Some("tag-1".to_string());
        assert!(summarize(&IRNode::ContentControl(control), &store).contains("tag=tag-1"));

        let mut start = BookmarkStart::new("b-1");
        start.name = Some("chapter".to_string());
        assert_eq!(
            summarize(&IRNode::BookmarkStart(start), &store),
            "id=b-1 name=chapter"
        );
        assert_eq!(
            summarize(&IRNode::BookmarkEnd(BookmarkEnd::new("b-1")), &store),
            "id=b-1"
        );

        let mut field = Field::new(Some("HYPERLINK".to_string()));
        field.runs.push(NodeId::new());
        assert!(summarize(&IRNode::Field(field), &store).contains("runs=1"));

        let mut rev = Revision::new(RevisionType::Insert);
        rev.content.push(NodeId::new());
        assert!(summarize(&IRNode::Revision(rev), &store).contains("content_nodes=1"));

        let mut theme = Theme::new();
        theme.name = Some("Office".to_string());
        theme.colors.push(ThemeColor {
            name: "accent1".to_string(),
            value: Some("FF0000".to_string()),
        });
        theme.fonts.major = Some("Calibri".to_string());
        assert!(summarize(&IRNode::Theme(theme), &store).contains("colors=1"));

        let media = MediaAsset::new("word/media/image1.png", MediaType::Image, 42);
        assert!(summarize(&IRNode::MediaAsset(media), &store).contains("size=42"));

        let mut custom = CustomXmlPart::new("customXml/item1.xml", 7);
        custom.root_element = Some("root".to_string());
        assert!(summarize(&IRNode::CustomXmlPart(custom), &store).contains("root=root"));

        let rel = RelationshipGraph::new("word/document.xml");
        assert!(summarize(&IRNode::RelationshipGraph(rel), &store).contains("rels=0"));

        let mut sig = DigitalSignature::new();
        sig.signature_id = Some("sig1".to_string());
        sig.signature_method = Some("rsa-sha256".to_string());
        assert!(summarize(&IRNode::DigitalSignature(sig), &store).contains("method=rsa-sha256"));

        let ext = ExtensionPart::new("word/unknown.bin", 9, ExtensionPartKind::Unknown);
        assert!(summarize(&IRNode::ExtensionPart(ext), &store).contains("size=9"));

        assert_eq!(
            summarize(&IRNode::StyleSet(StyleSet::new()), &store),
            "styles=0"
        );
        assert_eq!(
            summarize(&IRNode::NumberingSet(NumberingSet::new()), &store),
            "abstracts=0 nums=0"
        );

        let mut meta = DocumentMetadata::new();
        meta.title = Some("Doc".to_string());
        meta.creator = Some("Bob".to_string());
        assert_eq!(
            summarize(&IRNode::Metadata(meta), &store),
            "title=Doc author=Bob"
        );
    }

    #[test]
    fn summarize_covers_secondary_nodes_and_slide_style_signature() {
        let store = IrStore::new();

        assert_eq!(
            summarize(
                &IRNode::CommentExtensionSet(CommentExtensionSet::new()),
                &store
            ),
            "entries=0"
        );
        assert_eq!(
            summarize(&IRNode::CommentIdMap(CommentIdMap::new()), &store),
            "mappings=0"
        );
        assert_eq!(
            summarize(
                &IRNode::CommentRangeStart(CommentRangeStart::new("2")),
                &store
            ),
            "comment_id=2"
        );
        assert_eq!(
            summarize(&IRNode::CommentRangeEnd(CommentRangeEnd::new("2")), &store),
            "comment_id=2"
        );
        assert_eq!(
            summarize(
                &IRNode::CommentReference(CommentReference::new("2")),
                &store
            ),
            "comment_id=2"
        );
        assert_eq!(
            summarize(&IRNode::PeoplePart(PeoplePart::new()), &store),
            "people=0"
        );

        let smart = SmartArtPart {
            id: NodeId::new(),
            kind: "diagramData".to_string(),
            path: "ppt/diagrams/data1.xml".to_string(),
            root_element: None,
            point_count: None,
            connection_count: None,
            rel_ids: Vec::new(),
            span: None,
        };
        assert_eq!(
            summarize(&IRNode::SmartArtPart(smart), &store),
            "kind=diagramData path=ppt/diagrams/data1.xml"
        );

        let mut web = WebExtension::new();
        web.extension_id = Some("ext-id".to_string());
        web.store = Some("store".to_string());
        web.version = Some("1.0".to_string());
        assert!(summarize(&IRNode::WebExtension(web), &store).contains("properties=0"));

        let mut pane = WebExtensionTaskpane::new();
        pane.web_extension_ref = Some("ext-id".to_string());
        pane.dock_state = Some("right".to_string());
        pane.visibility = Some(true);
        assert_eq!(
            summarize(&IRNode::WebExtensionTaskpane(pane), &store),
            "ref=ext-id dock_state=right visible=true"
        );

        assert_eq!(
            summarize(&IRNode::GlossaryDocument(GlossaryDocument::new()), &store),
            "entries=0"
        );
        let mut glossary = GlossaryEntry::new();
        glossary.name = Some("quick".to_string());
        glossary.gallery = Some("auto".to_string());
        assert_eq!(
            summarize(&IRNode::GlossaryEntry(glossary), &store),
            "name=quick gallery=auto content_nodes=0"
        );

        let drawing = VmlDrawing::new("word/vmlDrawing1.vml");
        assert_eq!(
            summarize(&IRNode::VmlDrawing(drawing), &store),
            "path=word/vmlDrawing1.vml shapes=0"
        );
        let mut vml_shape = VmlShape::new();
        vml_shape.name = Some("shape1".to_string());
        vml_shape.rel_id = Some("rId1".to_string());
        vml_shape.image_target = Some("media/image1.png".to_string());
        assert_eq!(
            summarize(&IRNode::VmlShape(vml_shape), &store),
            "name=shape1 rel_id=rId1 image_target=media/image1.png"
        );

        let part = DrawingPart::new("word/drawings/drawing1.xml");
        assert_eq!(
            summarize(&IRNode::DrawingPart(part), &store),
            "path=word/drawings/drawing1.xml shapes=0"
        );

        let mut diag = Diagnostics::new();
        diag.entries.push(DiagnosticEntry {
            severity: DiagnosticSeverity::Warning,
            code: "W001".to_string(),
            message: "warn".to_string(),
            path: None,
        });
        assert_eq!(summarize(&IRNode::Diagnostics(diag), &store), "entries=1");

        let mut slide = docir_core::ir::Slide::new(1);
        slide.layout_id = Some("layout-1".to_string());
        slide.master_id = Some("master-1".to_string());
        assert_eq!(
            style_signature(&IRNode::Slide(slide), &store).unwrap(),
            "layout_id=layout-1 master_id=master-1"
        );
    }
}
