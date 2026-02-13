use super::*;
use crate::ooxml::part_utils::read_xml_part_and_rels;
use crate::zip_handler::PackageReader;

impl XlsxParser {
    pub(super) fn load_worksheet_drawings(
        &mut self,
        zip: &mut impl PackageReader,
        sheet_path: &str,
        relationships: &Relationships,
    ) -> Result<Vec<NodeId>, ParseError> {
        let mut drawings = Vec::new();
        for rel in relationships.get_by_type(rel_type::DRAWING) {
            let drawing_path = Relationships::resolve_target(sheet_path, &rel.target);
            if !zip.contains(&drawing_path) {
                continue;
            }
            let (drawing_xml, drawing_rels) = read_xml_part_and_rels(zip, &drawing_path)?;
            let drawing_id = self.parse_drawing(&drawing_xml, &drawing_path, &drawing_rels, zip)?;
            drawings.push(drawing_id);
        }
        Ok(drawings)
    }

    pub(super) fn load_worksheet_tables(
        &mut self,
        zip: &mut impl PackageReader,
        sheet_path: &str,
        relationships: &Relationships,
    ) -> Result<Vec<NodeId>, ParseError> {
        self.load_worksheet_parts(
            zip,
            sheet_path,
            relationships,
            rel_type::TABLE,
            |parser, zip, table_path, rel_id| parser.load_worksheet_table(zip, table_path, rel_id),
        )
    }

    pub(super) fn load_worksheet_pivots(
        &mut self,
        zip: &mut impl PackageReader,
        sheet_path: &str,
        relationships: &Relationships,
    ) -> Result<Vec<NodeId>, ParseError> {
        self.load_worksheet_parts(
            zip,
            sheet_path,
            relationships,
            rel_type::PIVOT_TABLE,
            |parser, zip, pivot_path, rel_id| parser.load_worksheet_pivot(zip, pivot_path, rel_id),
        )
    }

    pub(super) fn load_worksheet_comments(
        &mut self,
        zip: &mut impl PackageReader,
        sheet_path: &str,
        relationships: &Relationships,
        sheet_name: &str,
    ) -> Result<Vec<NodeId>, ParseError> {
        let mut comments = Vec::new();
        self.load_worksheet_comment_type(
            zip,
            sheet_path,
            relationships,
            sheet_name,
            rel_type::COMMENTS,
            parse_sheet_comments,
            &mut comments,
        )?;
        self.load_worksheet_comment_type(
            zip,
            sheet_path,
            relationships,
            sheet_name,
            rel_type::THREADED_COMMENTS,
            parse_threaded_comments,
            &mut comments,
        )?;
        Ok(comments)
    }

    fn load_worksheet_parts<R, F>(
        &mut self,
        zip: &mut R,
        sheet_path: &str,
        relationships: &Relationships,
        rel_type: &str,
        mut loader: F,
    ) -> Result<Vec<NodeId>, ParseError>
    where
        R: PackageReader,
        F: FnMut(&mut Self, &mut R, &str, &str) -> Result<NodeId, ParseError>,
    {
        let mut ids = Vec::new();
        for rel in relationships.get_by_type(rel_type) {
            let part_path = Relationships::resolve_target(sheet_path, &rel.target);
            if !zip.contains(&part_path) {
                continue;
            }
            let id = loader(self, zip, &part_path, &rel.id)?;
            ids.push(id);
        }
        Ok(ids)
    }

    fn load_worksheet_table(
        &mut self,
        zip: &mut impl PackageReader,
        table_path: &str,
        rel_id: &str,
    ) -> Result<NodeId, ParseError> {
        let table_xml = zip.read_file_string(table_path)?;
        let mut table = parse_table_definition(&table_xml, table_path)?;
        table.span = Some(SourceSpan::new(table_path).with_relationship(rel_id.to_string()));
        let id = table.id;
        self.store.insert(IRNode::TableDefinition(table));
        Ok(id)
    }

    fn load_worksheet_pivot(
        &mut self,
        zip: &mut impl PackageReader,
        pivot_path: &str,
        rel_id: &str,
    ) -> Result<NodeId, ParseError> {
        let pivot_xml = zip.read_file_string(pivot_path)?;
        let mut pivot = parse_pivot_table_definition(&pivot_xml, pivot_path)?;
        pivot.span = Some(SourceSpan::new(pivot_path).with_relationship(rel_id.to_string()));
        let id = pivot.id;
        self.store.insert(IRNode::PivotTable(pivot));
        Ok(id)
    }

    fn load_worksheet_comment_type(
        &mut self,
        zip: &mut impl PackageReader,
        sheet_path: &str,
        relationships: &Relationships,
        sheet_name: &str,
        rel_type: &str,
        parse_fn: fn(&str, &str, Option<&str>) -> Result<Vec<SheetComment>, ParseError>,
        out: &mut Vec<NodeId>,
    ) -> Result<(), ParseError> {
        for rel in relationships.get_by_type(rel_type) {
            let comments_path = Relationships::resolve_target(sheet_path, &rel.target);
            if !zip.contains(&comments_path) {
                continue;
            }
            let comments_xml = zip.read_file_string(&comments_path)?;
            let parsed = parse_fn(&comments_xml, &comments_path, Some(sheet_name))?;
            self.insert_sheet_comments(parsed, &comments_path, &rel.id, out);
        }
        Ok(())
    }

    fn insert_sheet_comments(
        &mut self,
        parsed: Vec<SheetComment>,
        comments_path: &str,
        rel_id: &str,
        out: &mut Vec<NodeId>,
    ) {
        for mut comment in parsed {
            comment.span =
                Some(SourceSpan::new(comments_path).with_relationship(rel_id.to_string()));
            let id = comment.id;
            self.store.insert(IRNode::SheetComment(comment));
            out.push(id);
        }
    }
}
