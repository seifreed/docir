mod tests {
    use crate::summary::{
        IRNode, abbreviate, cell_value_summary, content_signature, format_float, opt_bool, opt_u32,
        paragraph_text, short_hash, style_signature, summarize, summarize_cell, summarize_formula,
        summarize_paragraph, summarize_secondary, summarize_shape, text_from_paragraph,
    };
    use docir_core::ir::{
        Cell, CellFormula, CellValue, FontTable, FormulaType, Hyperlink, Paragraph, Run, Shape,
        ShapeText, ShapeTextParagraph, ShapeTextRun, ShapeType, Table, WebSettings, WordSettings,
        Worksheet,
    };
    use docir_core::security::ExternalRefType;
    use docir_core::types::NodeId;
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
        assert!(
            style_signature(&IRNode::Shape(shape), &store)
                .unwrap()
                .contains("has_text=true")
        );

        let secondary = summarize_secondary(&IRNode::CommentExtensionSet(
            docir_core::ir::CommentExtensionSet::new(),
        ));
        assert_eq!(secondary, "entries=0");
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
