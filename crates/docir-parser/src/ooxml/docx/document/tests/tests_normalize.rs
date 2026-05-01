use super::*;

#[test]
fn test_parse_comments_extended_and_ids() {
    let comments_extended_xml = r#"
        <w16cid:commentsEx xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
                           xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w16cid:commentExt w:id="1" w16cid:paraId="00AA" w16cid:parentParaId="00BB" w:done="1"/>
          <w16cid:commentExt w:id="2" w16cid:paraId="00CC" w:done="false"/>
        </w16cid:commentsEx>
        "#;
    let comments_ids_xml = r#"
        <w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
                            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w16cid:commentId w:id="1" w16cid:paraId="00AA" w16cid:parentParaId="00BB"/>
          <w16cid:commentId w:id="2" w16cid:paraId="00CC"/>
        </w16cid:commentsIds>
        "#;

    let mut parser = DocxParser::new();
    let extended_id = parser
        .parse_comments_extended(comments_extended_xml)
        .expect("comments extended");
    let ids_id = parser
        .parse_comments_ids(comments_ids_xml)
        .expect("comments ids");
    let store = parser.into_store();

    let ext = match store.get(extended_id) {
        Some(docir_core::ir::IRNode::CommentExtensionSet(set)) => set,
        _ => panic!("missing comment extension set"),
    };
    assert_eq!(ext.entries.len(), 2);
    assert_eq!(ext.entries[0].comment_id, "1");
    assert_eq!(ext.entries[0].para_id.as_deref(), Some("00AA"));
    assert_eq!(ext.entries[0].parent_para_id.as_deref(), Some("00BB"));
    assert_eq!(ext.entries[0].done, Some(true));
    assert_eq!(ext.entries[1].done, Some(false));

    let id_map = match store.get(ids_id) {
        Some(docir_core::ir::IRNode::CommentIdMap(map)) => map,
        _ => panic!("missing comment id map"),
    };
    assert_eq!(id_map.mappings.len(), 2);
    assert_eq!(id_map.mappings[0].comment_id, "1");
    assert_eq!(id_map.mappings[0].para_id.as_deref(), Some("00AA"));
    assert_eq!(id_map.mappings[1].comment_id, "2");
    assert_eq!(id_map.mappings[1].parent_para_id, None);
}

#[test]
fn test_parse_notes_endnote_kind_filters_footnotes() {
    let notes_xml = r#"
        <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:footnote w:id="1" w:type="separator">
            <w:p><w:r><w:t>Footnote ignored</w:t></w:r></w:p>
          </w:footnote>
          <w:endnote w:id="2" w:type="continuationSeparator">
            <w:p><w:r><w:t>Endnote kept</w:t></w:r></w:p>
          </w:endnote>
        </w:endnotes>
        "#;

    let mut parser = DocxParser::new();
    let ids = parser
        .parse_notes(notes_xml, NoteKind::Endnote, &Relationships::default())
        .expect("notes");
    assert_eq!(ids.len(), 1);
    let store = parser.into_store();
    let endnote = match store.get(ids[0]) {
        Some(docir_core::ir::IRNode::Endnote(n)) => n,
        _ => panic!("missing endnote"),
    };
    assert_eq!(endnote.endnote_id, "2");
    assert_eq!(endnote.note_type.as_deref(), Some("continuationSeparator"));
}

#[test]
fn test_parse_settings_entries_and_error_path() {
    let settings_xml = r#"
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:zoom w:percent="120"/>
          <w:defaultTabStop w:val="720"/>
        </w:settings>
        "#;

    let mut parser = DocxParser::new();
    let settings_id = parser.parse_settings(settings_xml).expect("settings");
    let web_settings_id = parser
        .parse_web_settings(settings_xml)
        .expect("web settings");
    let store = parser.into_store();

    let settings = match store.get(settings_id) {
        Some(docir_core::ir::IRNode::WordSettings(s)) => s,
        _ => panic!("missing settings"),
    };
    assert_eq!(settings.entries.len(), 3);
    assert_eq!(settings.entries[1].name, "w:zoom");
    assert!(settings.entries[1]
        .attributes
        .iter()
        .any(|a| a.name == "w:percent" && a.value == "120"));

    let web = match store.get(web_settings_id) {
        Some(docir_core::ir::IRNode::WebSettings(s)) => s,
        _ => panic!("missing web settings"),
    };
    assert_eq!(web.entries.len(), 3);

    let mut parser = DocxParser::new();
    let err = parser
        .parse_settings("<w:settings><")
        .expect_err("malformed settings should fail");
    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "word/settings.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_vml_picture_shape_and_style_lengths() {
    let xml = r#"
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:v="urn:schemas-microsoft-com:vml"
             xmlns:o="urn:schemas-microsoft-com:office:office"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:pict>
            <v:shape id="_x0000_s1025" name="VMLPic" o:title="Alt VML"
                     style="left:12pt;top:5mm;width:2cm;height:1in">
              <v:imagedata r:id="rIdImg"></v:imagedata>
            </v:shape>
          </w:pict>
        </w:r>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdImg"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="media/pict.png"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels");

    let mut reader = reader_from_str(xml);
    let mut parser = DocxParser::new();
    let mut run = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                run = Some(parse_run(&mut parser, &mut reader, &rels).expect("run"));
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }

    let run = run.expect("run parse");
    assert_eq!(run.embedded.len(), 1);
    let store = parser.into_store();
    let shape = match store.get(run.embedded[0]) {
        Some(docir_core::ir::IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    assert_eq!(shape.name.as_deref(), Some("VMLPic"));
    assert_eq!(shape.alt_text.as_deref(), Some("Alt VML"));
    assert_eq!(shape.media_target.as_deref(), Some("word/media/pict.png"));
    assert_eq!(shape.transform.x, 152400);
    assert_eq!(shape.transform.y, 180000);
    assert_eq!(shape.transform.width, 720000);
    assert_eq!(shape.transform.height, 914400);
}

#[test]
fn test_parse_bookmark_columns() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:bookmarkStart w:id="5" w:name="BM1" w:colFirst="1" w:colLast="3"/>
          <w:r><w:t>Text</w:t></w:r>
          <w:bookmarkEnd w:id="5"/>
        </w:p>
        "#;
    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let mut para = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                para = Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }
    let para = para.expect("para parsed");
    let store = parser.into_store();
    let para_node = match store.get(para.id) {
        Some(docir_core::ir::IRNode::Paragraph(p)) => p,
        _ => panic!("missing paragraph"),
    };
    let mut bookmark = None;
    for id in &para_node.runs {
        if let Some(docir_core::ir::IRNode::BookmarkStart(bm)) = store.get(*id) {
            bookmark = Some(bm);
            break;
        }
    }
    let bm = bookmark.expect("bookmark");
    assert_eq!(bm.name.as_deref(), Some("BM1"));
    assert_eq!(bm.col_first, Some(1));
    assert_eq!(bm.col_last, Some(3));
}

#[test]
fn test_parse_field_instruction_hyperlink() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:fldSimple w:instr="HYPERLINK &quot;https://example.com&quot; \\t &quot;_blank&quot;">
            <w:r><w:t>Link</w:t></w:r>
          </w:fldSimple>
        </w:p>
        "#;
    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let mut para = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                para = Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }
    let para = para.expect("para parsed");
    let store = parser.into_store();
    let para_node = match store.get(para.id) {
        Some(docir_core::ir::IRNode::Paragraph(p)) => p,
        _ => panic!("missing paragraph"),
    };
    let mut field = None;
    for id in &para_node.runs {
        if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id) {
            field = Some(f);
            break;
        }
    }
    let field = field.expect("field");
    let parsed = field.instruction_parsed.as_ref().expect("parsed");
    assert!(matches!(parsed.kind, docir_core::ir::FieldKind::Hyperlink));
    assert!(parsed.args.iter().any(|a| a == "https://example.com"));
    assert!(parsed.switches.iter().any(|s| s == "\\t"));
}
