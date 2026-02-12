use super::*;
use crate::diagnostics::attach_diagnostics_if_any;
use crate::zip_handler::PackageReader;

impl PptxParser {
    /// Parses the presentation and slides.
    pub fn parse_presentation(
        &mut self,
        zip: &mut impl PackageReader,
        presentation_xml: &str,
        presentation_rels: &Relationships,
        presentation_path: &str,
    ) -> Result<NodeId, ParseError> {
        let mut document = Document::new(DocumentFormat::Presentation);
        document.span = Some(SourceSpan::new(presentation_path));

        self.process_external_relationships(presentation_rels, presentation_path);

        let slide_refs = parse_slide_list(presentation_xml)?;

        if let Some(mut info) = parse_presentation_info(presentation_xml, presentation_path)? {
            info.span = Some(SourceSpan::new(presentation_path));
            let id = info.id;
            self.store.insert(IRNode::PresentationInfo(info));
            document.shared_parts.push(id);
        }

        for (index, rel_id) in slide_refs.into_iter().enumerate() {
            let rel = match presentation_rels.get(&rel_id) {
                Some(rel) => rel,
                None => {
                    push_warning(
                        &mut self.diagnostics,
                        "MISSING_RELATIONSHIP",
                        format!("Missing relationship for slide relId {}", rel_id),
                        Some(presentation_path),
                    );
                    continue;
                }
            };
            let slide_path = Relationships::resolve_target(presentation_path, &rel.target);

            let (slide_xml, slide_rels) = read_xml_part_and_rels(zip, &slide_path)?;

            self.process_external_relationships(&slide_rels, &slide_path);

            let mut notes_slide_id: Option<NodeId> = None;
            let notes_text = if let Some(rel) = slide_rels.get_first_by_type(rel_type::NOTES_SLIDE)
            {
                let notes_path = Relationships::resolve_target(&slide_path, &rel.target);
                if let Some((notes_xml, notes_rels)) =
                    read_xml_part_and_rels_optional(zip, &notes_path)?
                {
                    let (notes_node, notes_text) =
                        parse_notes_slide(&notes_xml, &notes_path, &notes_rels, self, zip)?;
                    let notes_id = notes_node.id;
                    self.store.insert(IRNode::NotesSlide(notes_node));
                    notes_slide_id = Some(notes_id);
                    Some(notes_text)
                } else {
                    None
                }
            } else {
                None
            };

            let slide_id = self.parse_slide(
                zip,
                &slide_xml,
                (index + 1) as u32,
                &slide_path,
                &slide_rels,
                notes_text.as_deref(),
                notes_slide_id,
            )?;
            document.content.push(slide_id);
        }

        self.load_presentation_parts(&mut document, zip)?;

        // Parse slide masters and layouts
        for rel in presentation_rels.get_by_type(rel_type::SLIDE_MASTER) {
            let master_path = Relationships::resolve_target(presentation_path, &rel.target);
            if !zip.contains(&master_path) {
                continue;
            }
            let (master_xml, master_rels) = read_xml_part_and_rels(zip, &master_path)?;

            let master_name = extract_c_sld_name(&master_xml);
            let master_shapes =
                self.parse_shapes_from_xml(&master_xml, &master_path, &master_rels, zip)?;
            let master_meta = parse_slide_master_meta(&master_xml, &master_path)?;

            let mut layout_ids = Vec::new();
            for layout_rel in master_rels.get_by_type(rel_type::SLIDE_LAYOUT) {
                let layout_path = Relationships::resolve_target(&master_path, &layout_rel.target);
                if !zip.contains(&layout_path) {
                    continue;
                }
                let layout_xml = zip.read_file_string(&layout_path)?;
                let layout_id =
                    self.parse_slide_layout(&layout_xml, &layout_path, &master_rels, zip)?;
                layout_ids.push(layout_id);
            }

            let mut master = docir_core::ir::SlideMaster::new();
            master.name = master_name;
            master.preserve = master_meta.preserve;
            master.show_master_sp = master_meta.show_master_sp;
            master.show_master_ph_anim = master_meta.show_master_ph_anim;
            master.shapes = master_shapes;
            master.layouts = layout_ids.clone();
            master.span = Some(SourceSpan::new(&master_path));
            let master_id = master.id;
            self.store.insert(IRNode::SlideMaster(master));

            document.shared_parts.push(master_id);
            document.shared_parts.extend(layout_ids);
        }

        // Parse notes master
        if let Some(rel) = presentation_rels.get_first_by_type(rel_type::NOTES_MASTER) {
            let notes_master_path = Relationships::resolve_target(presentation_path, &rel.target);
            if zip.contains(&notes_master_path) {
                let notes_master_xml = zip.read_file_string(&notes_master_path)?;
                let mut notes_master = docir_core::ir::NotesMaster::new();
                notes_master.name = extract_c_sld_name(&notes_master_xml);
                notes_master.shapes = self.parse_shapes_from_xml(
                    &notes_master_xml,
                    &notes_master_path,
                    presentation_rels,
                    zip,
                )?;
                notes_master.span = Some(SourceSpan::new(&notes_master_path));
                let id = notes_master.id;
                self.store.insert(IRNode::NotesMaster(notes_master));
                document.shared_parts.push(id);
            }
        }

        // Parse handout master
        if let Some(rel) = presentation_rels.get_first_by_type(rel_type::HANDOUT_MASTER) {
            let handout_path = Relationships::resolve_target(presentation_path, &rel.target);
            if zip.contains(&handout_path) {
                let handout_xml = zip.read_file_string(&handout_path)?;
                let mut handout = docir_core::ir::HandoutMaster::new();
                handout.name = extract_c_sld_name(&handout_xml);
                handout.shapes = self.parse_shapes_from_xml(
                    &handout_xml,
                    &handout_path,
                    presentation_rels,
                    zip,
                )?;
                handout.span = Some(SourceSpan::new(&handout_path));
                let id = handout.id;
                self.store.insert(IRNode::HandoutMaster(handout));
                document.shared_parts.push(id);
            }
        }

        document.shared_parts.extend(self.chart_nodes.drain(..));
        document.security = std::mem::take(&mut self.security_info);

        let mut diagnostics = std::mem::replace(&mut self.diagnostics, Diagnostics::new());
        if !diagnostics.entries.is_empty() {
            diagnostics.span = Some(SourceSpan::new(presentation_path));
            attach_diagnostics_if_any(&mut self.store, &mut document, diagnostics);
        }

        let doc_id = document.id;
        self.store.insert(IRNode::Document(document));
        Ok(doc_id)
    }
}
