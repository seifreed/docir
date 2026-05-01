mod tests {
    use crate::summary::{
        abbreviate, cell_value_summary, content_signature, format_float, opt_bool, opt_u32,
        paragraph_text, short_hash, style_signature, summarize, summarize_cell, summarize_formula,
        summarize_paragraph, summarize_primary, summarize_secondary, summarize_shape,
        text_from_paragraph, IRNode,
    };
    use docir_core::ir::{
        BookmarkEnd, BookmarkStart, Cell, CellFormula, CellValue, Comment, ContentControl,
        CustomXmlPart, DiagnosticEntry, DiagnosticSeverity, Diagnostics, DigitalSignature,
        DocumentMetadata, DrawingPart, Endnote, ExtensionPart, ExtensionPartKind, Field, FontTable,
        Footer, FormulaType, GlossaryDocument, GlossaryEntry, Header, Hyperlink, NumberingSet,
        Paragraph, PeoplePart, RelationshipGraph, Revision, RevisionType, Run, Shape, ShapeText,
        ShapeTextParagraph, ShapeTextRun, ShapeType, SmartArtPart, StyleSet, Table, TableCell,
        TableRow, Theme, ThemeColor, VmlDrawing, VmlShape, WebExtension, WebExtensionTaskpane,
        WebSettings, WordSettings, Worksheet,
    };
    use docir_core::ir::{CommentExtensionSet, CommentIdMap, CommentRangeEnd, CommentRangeStart};
    use docir_core::ir::{CommentReference, MediaAsset, MediaType};
    use docir_core::security::ExternalRefType;
    use docir_core::types::{DocumentFormat, NodeId};
    use docir_core::visitor::IrStore;

    #[test]
    fn helper_functions_are_deterministic() {
        assert_eq!(opt_bool(Some(true)), "true");
        assert_eq!(opt_bool(Some(false)), "false");
        assert_eq!(opt_bool(None), "-");
        assert_eq!(opt_u32(Some(42)), "42");
        assert_eq!(opt_u32(None), "-");
        assert_eq!(abbreviate("abc", 3), "abc");
        assert_eq!(abbreviate("abcdef", 3), "abc...");
        assert_eq!(format_float(2.0), "2");
        assert_eq!(format_float(2.5), "2.500000");
        assert_eq!(short_hash("same"), short_hash("same"));
    }

    #[test]
    fn summarizes_primary_word_and_security_nodes() {
        let store = IrStore::new();

        let doc = IRNode::Document(docir_core::ir::Document::new(
            DocumentFormat::WordProcessing,
        ));
        assert!(summarize_primary(&doc, &store)
            .unwrap()
            .contains("format=WordProcessing"));

        let mut section = docir_core::ir::Section::new();
        section.name = Some("Body".to_string());
        assert!(summarize_primary(&IRNode::Section(section), &store)
            .unwrap()
            .contains("name=Body"));

        let para_id = NodeId::new();
        let mut with_store = IrStore::new();
        let run = Run::new("Hello");
        with_store.insert(IRNode::Run(run.clone()));
        let mut para = Paragraph::new();
        para.id = para_id;
        para.runs.push(run.id);
        assert!(summarize_primary(&IRNode::Paragraph(para), &with_store)
            .unwrap()
            .contains("text=\"Hello\""));

        let run = Run::new("Run text");
        assert!(summarize_primary(&IRNode::Run(run.clone()), &store)
            .unwrap()
            .contains("bold=-"));

        let mut link = Hyperlink::new("https://example.test", true);
        link.runs.push(run.id);
        assert!(summarize_primary(&IRNode::Hyperlink(link), &with_store)
            .unwrap()
            .contains("external=true"));

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
        assert!(summarize_primary(&IRNode::ExternalReference(ext), &store)
            .unwrap()
            .contains("target=http://example.test"));
    }

    #[test]
    fn secondary_summary_and_signatures_cover_key_branches() {
        let mut store = IrStore::new();
        let run_a = Run::new("A");
        let run_b = Run::new("B");
        store.insert(IRNode::Run(run_a.clone()));
        store.insert(IRNode::Run(run_b.clone()));

        let mut para = Paragraph::new();
        para.runs = vec![run_a.id, run_b.id];
        let para_sig = content_signature(&IRNode::Paragraph(para), &store).unwrap();
        assert_eq!(para_sig, "AB");

        let mut cell = Cell::new("C3", 2, 2);
        cell.value = CellValue::Number(9.0);
        cell.formula = Some(CellFormula {
            text: "SUM(A1:A2)".to_string(),
            formula_type: FormulaType::Normal,
            shared_index: None,
            shared_ref: None,
            is_array: false,
            array_ref: None,
        });
        let cell_sig = content_signature(&IRNode::Cell(cell.clone()), &store).unwrap();
        assert!(cell_sig.contains("C3=n:9;SUM(A1:A2)"));

        let mut worksheet = Worksheet::new("Data", 1);
        let cell_id = NodeId::new();
        store.insert(IRNode::Cell(cell));
        worksheet.cells.push(cell_id);
        let ws_sig = content_signature(&IRNode::Worksheet(worksheet), &store).unwrap();
        assert_eq!(ws_sig.len(), 16);

        let shape = Shape {
            text: Some(ShapeText {
                paragraphs: vec![ShapeTextParagraph {
                    runs: vec![ShapeTextRun {
                        text: "Title".to_string(),
                        bold: None,
                        italic: None,
                        font_size: None,
                        font_family: None,
                    }],
                    alignment: None,
                }],
            }),
            ..Shape::new(ShapeType::TextBox)
        };
        assert_eq!(
            content_signature(&IRNode::Shape(shape.clone()), &store).unwrap(),
            "Title"
        );
        assert!(style_signature(&IRNode::Shape(shape), &store)
            .unwrap()
            .contains("has_text=true"));

        let secondary = summarize_secondary(&IRNode::CommentExtensionSet(
            docir_core::ir::CommentExtensionSet::new(),
        ));
        assert_eq!(secondary, "entries=0");
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

    #[test]
    fn summary_and_signature_cover_macro_and_hyperlink_fallback_paths() {
        let mut store = IrStore::new();

        let run_a = Run::new("PartA");
        let run_b = Run::new("PartB");
        store.insert(IRNode::Run(run_a.clone()));
        store.insert(IRNode::Run(run_b.clone()));

        let mut link = Hyperlink::new("https://example.test", true);
        link.runs = vec![run_a.id, run_b.id];
        let link_id = link.id;
        store.insert(IRNode::Hyperlink(link.clone()));

        let mut para = Paragraph::new();
        para.runs.push(link_id);
        let para_text = summarize_paragraph(&para, &store);
        assert!(para_text.contains("PartAPartB"));

        assert_eq!(
            content_signature(&IRNode::Hyperlink(link), &store).as_deref(),
            Some("https://example.test")
        );

        let mut module = docir_core::security::MacroModule::new(
            "AutoOpen",
            docir_core::security::MacroModuleType::Standard,
        );
        module
            .suspicious_calls
            .push(docir_core::security::SuspiciousCall {
                name: "Shell".to_string(),
                category: docir_core::security::SuspiciousCallCategory::ShellExecution,
                line: Some(1),
            });
        let module_summary = summarize(&IRNode::MacroModule(module.clone()), &store);
        assert!(module_summary.contains("suspicious_calls=1"));
        assert_eq!(
            content_signature(&IRNode::MacroModule(module), &store).as_deref(),
            Some("AutoOpen")
        );

        let mut project = docir_core::security::MacroProject::new();
        project.name = Some("VBAProject".to_string());
        assert_eq!(
            content_signature(&IRNode::MacroProject(project), &store).as_deref(),
            Some("VBAProject")
        );

        let mut activex = docir_core::ActiveXControl::new();
        activex.name = Some("Btn".to_string());
        assert_eq!(
            content_signature(&IRNode::ActiveXControl(activex), &store).as_deref(),
            Some("Btn")
        );

        let mut ole = docir_core::security::OleObject::new();
        ole.name = Some("Object1".to_string());
        assert_eq!(
            content_signature(&IRNode::OleObject(ole), &store).as_deref(),
            Some("Object1")
        );

        let defined = docir_core::ir::DefinedName {
            id: NodeId::new(),
            name: "MyName".to_string(),
            value: "Sheet1!$A$1".to_string(),
            local_sheet_id: None,
            hidden: false,
            comment: None,
            span: None,
        };
        assert_eq!(
            content_signature(&IRNode::DefinedName(defined), &store).as_deref(),
            Some("MyName")
        );

        let table = docir_core::ir::TableDefinition {
            id: NodeId::new(),
            name: Some("TableFallback".to_string()),
            display_name: None,
            ref_range: None,
            header_row_count: None,
            totals_row_count: None,
            columns: vec![],
            span: None,
        };
        assert_eq!(
            content_signature(&IRNode::TableDefinition(table), &store).as_deref(),
            Some("TableFallback")
        );
    }

    #[test]
    fn style_signature_and_shape_text_cover_remaining_branches() {
        let store = IrStore::new();

        let para = Paragraph::new();
        let para_sig = style_signature(&IRNode::Paragraph(para), &store).unwrap();
        assert!(para_sig.starts_with('{'));

        let run_sig = style_signature(&IRNode::Run(Run::new("r")), &store).unwrap();
        assert!(run_sig.starts_with('{'));

        let table_sig = style_signature(&IRNode::Table(Table::new()), &store).unwrap();
        assert!(table_sig.starts_with('{'));

        let shape = Shape {
            text: Some(ShapeText {
                paragraphs: vec![
                    ShapeTextParagraph {
                        runs: vec![ShapeTextRun {
                            text: "L1".to_string(),
                            bold: None,
                            italic: None,
                            font_size: None,
                            font_family: None,
                        }],
                        alignment: None,
                    },
                    ShapeTextParagraph {
                        runs: vec![ShapeTextRun {
                            text: "L2".to_string(),
                            bold: None,
                            italic: None,
                            font_size: None,
                            font_family: None,
                        }],
                        alignment: None,
                    },
                ],
            }),
            ..Shape::new(ShapeType::TextBox)
        };
        let summary = summarize_shape(&shape);
        assert!(summary.contains("L1\nL2"));
    }

    #[test]
    fn summary_covers_remaining_primary_and_presentation_paths() {
        let store = IrStore::new();

        let mut slide = docir_core::ir::Slide::new(2);
        slide.name = Some("Deck".to_string());
        assert!(summarize(&IRNode::Slide(slide), &store).contains("name=Deck"));

        let mut activex = docir_core::ActiveXControl::new();
        activex.name = Some("Button1".to_string());
        activex.clsid = Some("{clsid}".to_string());
        activex.prog_id = Some("Forms.CommandButton.1".to_string());
        let activex_summary = summarize(&IRNode::ActiveXControl(activex.clone()), &store);
        assert!(activex_summary.contains("name=Button1"));
        assert_eq!(
            content_signature(&IRNode::ActiveXControl(activex), &store).as_deref(),
            Some("Forms.CommandButton.1")
        );

        assert_eq!(
            summarize(&IRNode::WordSettings(WordSettings::new()), &store),
            "entries=0"
        );
        assert_eq!(
            summarize(&IRNode::WebSettings(WebSettings::new()), &store),
            "entries=0"
        );
        assert_eq!(
            summarize(&IRNode::FontTable(FontTable::new()), &store),
            "fonts=0"
        );

        let mut ext = docir_core::security::ExternalReference::new(
            ExternalRefType::Hyperlink,
            "https://example.test/ext",
        );
        ext.target = "https://example.test/ext".to_string();
        assert_eq!(
            content_signature(&IRNode::ExternalReference(ext), &store).as_deref(),
            Some("https://example.test/ext")
        );

        let mut ole = docir_core::security::OleObject::new();
        ole.name = Some("object".to_string());
        ole.prog_id = Some("Excel.Sheet".to_string());
        ole.data_hash = Some("deadbeef".to_string());
        ole.is_linked = true;
        ole.size_bytes = 128;
        let ole_summary = summarize(&IRNode::OleObject(ole), &store);
        assert!(ole_summary.contains("prog_id=Excel.Sheet"));
        assert!(ole_summary.contains("hash=deadbeef"));
    }

    #[test]
    fn helper_paths_cover_cell_variants_and_non_run_paragraph_nodes() {
        let mut store = IrStore::new();

        assert_eq!(cell_value_summary(&CellValue::Empty), "empty");
        assert_eq!(cell_value_summary(&CellValue::Boolean(true)), "b:true");
        assert_eq!(
            cell_value_summary(&CellValue::InlineString("x".to_string())),
            "is:x"
        );
        assert_eq!(cell_value_summary(&CellValue::SharedString(3)), "ss:3");
        assert_eq!(cell_value_summary(&CellValue::DateTime(1.25)), "dt:1.25");
        assert!(
            cell_value_summary(&CellValue::Error(docir_core::ir::CellError::Ref)).contains("Ref")
        );

        let mut cell = Cell::new("B2", 1, 1);
        cell.value = CellValue::Boolean(false);
        cell.formula = Some(CellFormula {
            text: "A1+1".to_string(),
            formula_type: FormulaType::Array,
            shared_index: None,
            shared_ref: None,
            is_array: true,
            array_ref: Some("B2:B3".to_string()),
        });
        let cell_summary = summarize_cell(&cell);
        assert!(cell_summary.contains("value=bool:false"));
        assert!(cell_summary.contains("type=Array"));
        assert_eq!(
            summarize_formula(cell.formula.as_ref().expect("formula")),
            "A1+1 type=Array"
        );

        let text_cell_id = NodeId::new();
        store.insert(IRNode::Cell(cell.clone()));
        let mut ws = Worksheet::new("SheetX", 7);
        ws.cells.push(text_cell_id);
        ws.cells.push(NodeId::new());
        let ws_sig =
            content_signature(&IRNode::Worksheet(ws), &store).expect("worksheet signature");
        assert_eq!(ws_sig.len(), 16);

        let run = Run::new("text");
        let run_id = run.id;
        store.insert(IRNode::Run(run));
        let non_text_node_id = NodeId::new();
        store.insert(IRNode::Cell(Cell::new("C3", 2, 2)));
        let mut para = Paragraph::new();
        para.runs = vec![run_id, non_text_node_id];
        assert_eq!(paragraph_text(&para, &store), "text");
        assert_eq!(text_from_paragraph(&para, &store), "text");
    }

    #[test]
    fn summarize_secondary_returns_placeholder_for_unsupported_node() {
        let run = Run::new("unsupported");
        assert_eq!(summarize_secondary(&IRNode::Run(run)), "unsupported=Run",);
    }
}
