use docir_core::ir::IRNode;

pub(super) fn intrinsic_key(node: &IRNode) -> Option<String> {
    intrinsic_key_spreadsheet(node)
        .or_else(|| intrinsic_key_word(node))
        .or_else(|| intrinsic_key_presentation(node))
        .or_else(|| intrinsic_key_security(node))
}

fn intrinsic_key_spreadsheet(node: &IRNode) -> Option<String> {
    match node {
        IRNode::Worksheet(ws) => Some(format!("Worksheet[{}]", ws.name)),
        IRNode::Cell(cell) => Some(format!("Cell[{}]", cell.reference)),
        IRNode::DefinedName(def) => Some(format!("DefinedName[{}]", def.name)),
        IRNode::TableDefinition(table) => table
            .display_name
            .as_ref()
            .map(|name| format!("TableDefinition[{name}]")),
        IRNode::PivotTable(pivot) => pivot
            .name
            .as_ref()
            .map(|name| format!("PivotTable[{name}]")),
        IRNode::WorksheetDrawing(_) => Some("WorksheetDrawing".to_string()),
        IRNode::CalcChain(_) => Some("CalcChain".to_string()),
        IRNode::SheetComment(comment) => Some(format!("SheetComment[{}]", comment.cell_ref)),
        IRNode::SheetMetadata(_) => Some("SheetMetadata".to_string()),
        IRNode::ExternalLinkPart(part) => Some(format!(
            "ExternalLinkPart[{}]",
            part.target.as_deref().unwrap_or("-")
        )),
        IRNode::ConnectionPart(part) => Some(format!("ConnectionPart[{}]", part.entries.len())),
        IRNode::SlicerPart(part) => Some(format!(
            "SlicerPart[{}]",
            part.name.as_deref().unwrap_or("-")
        )),
        IRNode::TimelinePart(part) => Some(format!(
            "TimelinePart[{}]",
            part.name.as_deref().unwrap_or("-")
        )),
        IRNode::QueryTablePart(part) => Some(format!(
            "QueryTablePart[{}]",
            part.name.as_deref().unwrap_or("-")
        )),
        IRNode::SharedStringTable(_) => Some("SharedStringTable".to_string()),
        IRNode::SpreadsheetStyles(_) => Some("SpreadsheetStyles".to_string()),
        IRNode::ConditionalFormat(_) => Some("ConditionalFormat".to_string()),
        IRNode::DataValidation(_) => Some("DataValidation".to_string()),
        IRNode::PivotCache(cache) => Some(format!("PivotCache[{}]", cache.cache_id)),
        IRNode::PivotCacheRecords(records) => Some(format!(
            "PivotCacheRecords[{}]",
            records.record_count.unwrap_or(0)
        )),
        IRNode::WorkbookProperties(_) => Some("WorkbookProperties".to_string()),
        _ => None,
    }
}

fn intrinsic_key_word(node: &IRNode) -> Option<String> {
    match node {
        IRNode::Hyperlink(link) => Some(format!("Hyperlink[{}]", link.target)),
        IRNode::StyleSet(styles) => Some(format!("StyleSet[{}]", styles.styles.len())),
        IRNode::NumberingSet(nums) => Some(format!("NumberingSet[{}]", nums.abstract_nums.len())),
        IRNode::Comment(comment) => Some(format!("Comment[{}]", comment.comment_id)),
        IRNode::Footnote(note) => Some(format!("Footnote[{}]", note.footnote_id)),
        IRNode::Endnote(note) => Some(format!("Endnote[{}]", note.endnote_id)),
        IRNode::Header(_) => Some("Header".to_string()),
        IRNode::Footer(_) => Some("Footer".to_string()),
        IRNode::WordSettings(_) => Some("WordSettings".to_string()),
        IRNode::WebSettings(_) => Some("WebSettings".to_string()),
        IRNode::FontTable(_) => Some("FontTable".to_string()),
        IRNode::ContentControl(_) => Some("ContentControl".to_string()),
        IRNode::BookmarkStart(start) => Some(format!("BookmarkStart[{}]", start.bookmark_id)),
        IRNode::BookmarkEnd(end) => Some(format!("BookmarkEnd[{}]", end.bookmark_id)),
        IRNode::Field(field) => Some(format!(
            "Field[{}]",
            field.instruction.as_deref().unwrap_or("-")
        )),
        IRNode::Revision(rev) => Some(format!("Revision[{:?}]", rev.change_type)),
        IRNode::CommentExtensionSet(_) => Some("CommentExtensionSet".to_string()),
        IRNode::CommentIdMap(_) => Some("CommentIdMap".to_string()),
        IRNode::CommentRangeStart(start) => {
            Some(format!("CommentRangeStart[{}]", start.comment_id))
        }
        IRNode::CommentRangeEnd(end) => Some(format!("CommentRangeEnd[{}]", end.comment_id)),
        IRNode::CommentReference(reference) => {
            Some(format!("CommentReference[{}]", reference.comment_id))
        }
        IRNode::GlossaryDocument(_) => Some("GlossaryDocument".to_string()),
        IRNode::GlossaryEntry(entry) => Some(format!(
            "GlossaryEntry[{}]",
            entry.name.as_deref().unwrap_or("-")
        )),
        IRNode::VmlDrawing(drawing) => Some(format!("VmlDrawing[{}]", drawing.path)),
        IRNode::VmlShape(shape) => Some(format!(
            "VmlShape[{}]",
            shape.name.as_deref().unwrap_or("-")
        )),
        IRNode::DrawingPart(part) => Some(format!("DrawingPart[{}]", part.path)),
        _ => None,
    }
}

fn intrinsic_key_presentation(node: &IRNode) -> Option<String> {
    match node {
        IRNode::Slide(slide) => Some(format!("Slide[{}]", slide.number)),
        IRNode::Shape(shape) => shape.name.as_ref().map(|name| format!("Shape[{name}]")),
        IRNode::SlideMaster(_) => Some("SlideMaster".to_string()),
        IRNode::SlideLayout(_) => Some("SlideLayout".to_string()),
        IRNode::NotesMaster(_) => Some("NotesMaster".to_string()),
        IRNode::HandoutMaster(_) => Some("HandoutMaster".to_string()),
        IRNode::NotesSlide(_) => Some("NotesSlide".to_string()),
        IRNode::ChartData(_) => Some("ChartData".to_string()),
        IRNode::PresentationProperties(_) => Some("PresentationProperties".to_string()),
        IRNode::ViewProperties(_) => Some("ViewProperties".to_string()),
        IRNode::TableStyleSet(_) => Some("TableStyleSet".to_string()),
        IRNode::PptxCommentAuthor(author) => {
            Some(format!("PptxCommentAuthor[{}]", author.author_id))
        }
        IRNode::PptxComment(_) => Some("PptxComment".to_string()),
        IRNode::PresentationTag(tag) => Some(format!("PresentationTag[{}]", tag.name)),
        IRNode::PresentationInfo(_) => Some("PresentationInfo".to_string()),
        IRNode::PeoplePart(_) => Some("PeoplePart".to_string()),
        IRNode::SmartArtPart(part) => Some(format!("SmartArtPart[{}]", part.kind)),
        _ => None,
    }
}

fn intrinsic_key_security(node: &IRNode) -> Option<String> {
    match node {
        IRNode::ExternalReference(ext) => Some(format!("ExternalReference[{}]", ext.target)),
        IRNode::ActiveXControl(ctrl) => ctrl
            .prog_id
            .as_ref()
            .map(|p| format!("ActiveXControl[{p}]")),
        IRNode::MacroModule(module) => Some(format!("MacroModule[{}]", module.name)),
        IRNode::MacroProject(project) => project
            .name
            .as_ref()
            .map(|name| format!("MacroProject[{name}]")),
        IRNode::OleObject(ole) => {
            if let Some(name) = &ole.name {
                Some(format!("OleObject[{}]", name))
            } else {
                ole.prog_id
                    .as_ref()
                    .map(|prog_id| format!("OleObject[{prog_id}]"))
            }
        }
        IRNode::WebExtension(ext) => Some(format!(
            "WebExtension[{}]",
            ext.extension_id.as_deref().unwrap_or("-")
        )),
        IRNode::WebExtensionTaskpane(pane) => Some(format!(
            "WebExtensionTaskpane[{}]",
            pane.web_extension_ref.as_deref().unwrap_or("-")
        )),
        IRNode::Diagnostics(_) => Some("Diagnostics".to_string()),
        _ => None,
    }
}
