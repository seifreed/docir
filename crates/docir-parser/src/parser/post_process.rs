use super::*;
use crate::diagnostics::attach_diagnostics_if_any_to_store;
use crate::zip_handler::PackageReader;

impl OoxmlParser {
    pub(super) fn post_process_ooxml(
        &self,
        zip: &mut impl PackageReader,
        content_types: &ContentTypes,
        store: &mut IrStore,
        root_id: NodeId,
        metrics: &mut Option<ParseMetrics>,
    ) -> Result<(), ParseError> {
        // Link shapes/animations to shared parts (charts, SmartArt, media, OLE)
        self.link_shapes_to_shared_parts(store);

        let start = std::time::Instant::now();
        self.add_extension_parts_and_diagnostics(zip, content_types, store, root_id)?;
        if let Some(m) = metrics.as_mut() {
            m.extension_parts_ms = start.elapsed().as_millis();
        }

        let start = std::time::Instant::now();
        normalize_store(store, root_id);
        if let Some(m) = metrics.as_mut() {
            m.normalization_ms = start.elapsed().as_millis();
        }

        Ok(())
    }

    fn add_extension_parts_and_diagnostics(
        &self,
        zip: &mut impl PackageReader,
        content_types: &ContentTypes,
        store: &mut IrStore,
        root_id: NodeId,
    ) -> Result<(), ParseError> {
        let mut seen_paths = coverage::collect_seen_paths(store);
        seen_paths.insert("[Content_Types].xml".to_string());
        if let Some(IRNode::Document(doc)) = store.get(root_id) {
            if doc.metadata.is_some() {
                if zip.contains("docProps/core.xml") {
                    seen_paths.insert("docProps/core.xml".to_string());
                }
                if zip.contains("docProps/app.xml") {
                    seen_paths.insert("docProps/app.xml".to_string());
                }
                if zip.contains("docProps/custom.xml") {
                    seen_paths.insert("docProps/custom.xml".to_string());
                }
            }
            match doc.format {
                DocumentFormat::WordProcessing => {
                    if zip.contains("word/comments.xml") {
                        seen_paths.insert("word/comments.xml".to_string());
                    }
                }
                DocumentFormat::Presentation => {
                    if zip.contains("ppt/commentAuthors.xml") {
                        seen_paths.insert("ppt/commentAuthors.xml".to_string());
                    }
                }
                _ => {}
            }
        }

        let mut extension_ids = Vec::new();
        let mut diagnostics = Diagnostics::new();
        diagnostics.span = Some(SourceSpan::new("package"));
        let mut unparsed_count: usize = 0;

        let all_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| !p.ends_with('/') && !p.starts_with("[trash]/"))
            .collect();
        for path in all_paths {
            if seen_paths.contains(&path) {
                continue;
            }
            let data = zip.read_file(&path)?;
            let mut part = ExtensionPart::new(&path, data.len() as u64, ExtensionPartKind::Unknown);
            part.content_type = content_types.get_content_type(&path).map(|s| s.to_string());
            part.span = Some(SourceSpan::new(&path));
            let part_id = part.id;
            store.insert(IRNode::ExtensionPart(part));
            extension_ids.push(part_id);
            unparsed_count += 1;

            push_warning(
                &mut diagnostics,
                "UNPARSED_PART",
                format!("No parser registered for part: {}", path),
                Some(&path),
            );
        }

        let format = match store.get(root_id) {
            Some(IRNode::Document(doc)) => doc.format,
            _ => DocumentFormat::WordProcessing,
        };
        let coverage_entries = coverage::build_coverage_diagnostics(
            format,
            content_types,
            zip.file_names(),
            &seen_paths,
        );
        diagnostics.entries.extend(coverage_entries);
        push_info(
            &mut diagnostics,
            "UNPARSED_SUMMARY",
            format!("unparsed parts: {}", unparsed_count),
            None,
        );

        if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
            doc.shared_parts.extend(extension_ids);
        }

        if !diagnostics.entries.is_empty() {
            attach_diagnostics_if_any_to_store(store, root_id, diagnostics);
        }

        Ok(())
    }

    fn link_shapes_to_shared_parts(&self, store: &mut IrStore) {
        use docir_core::ir::IRNode;
        use docir_core::types::NodeType;

        let mut chart_by_path: HashMap<String, NodeId> = HashMap::new();
        let mut smartart_by_path: HashMap<String, NodeId> = HashMap::new();
        let mut media_by_path: HashMap<String, NodeId> = HashMap::new();
        let mut ole_by_path: HashMap<String, NodeId> = HashMap::new();

        for (id, node) in store.iter() {
            match node {
                IRNode::ChartData(chart) => {
                    if let Some(span) = chart.span.as_ref() {
                        chart_by_path.insert(span.file_path.clone(), *id);
                    }
                }
                IRNode::SmartArtPart(part) => {
                    smartart_by_path.insert(part.path.clone(), *id);
                }
                IRNode::MediaAsset(media) => {
                    media_by_path.insert(media.path.clone(), *id);
                }
                IRNode::OleObject(ole) => {
                    if let Some(span) = ole.span.as_ref() {
                        ole_by_path.insert(span.file_path.clone(), *id);
                    }
                }
                _ => {}
            }
        }

        let shape_ids: Vec<NodeId> = store.iter_ids_by_type(NodeType::Shape).collect();
        for shape_id in shape_ids {
            let Some(IRNode::Shape(shape)) = store.get_mut(shape_id) else {
                continue;
            };
            if let Some(target) = shape.media_target.as_ref() {
                if let Some(id) = utils::resolve_media_asset(&media_by_path, target) {
                    shape.media_asset = Some(id);
                }
                if let Some(id) = chart_by_path.get(target) {
                    shape.chart_id = Some(*id);
                }
                if let Some(id) = ole_by_path.get(target) {
                    shape.ole_object = Some(*id);
                    shape.shape_type = docir_core::ir::ShapeType::OleObject;
                }
            }
            if !shape.related_targets.is_empty() {
                for target in &shape.related_targets {
                    if let Some(id) = smartart_by_path.get(target) {
                        shape.smartart_parts.push(*id);
                    }
                    if let Some(id) = ole_by_path.get(target) {
                        shape.ole_object = Some(*id);
                        shape.shape_type = docir_core::ir::ShapeType::OleObject;
                    }
                }
                shape.smartart_parts.sort_by_key(|id| id.as_u64());
                shape.smartart_parts.dedup();
            }
        }

        let slide_ids: Vec<NodeId> = store.iter_ids_by_type(NodeType::Slide).collect();
        for slide_id in slide_ids {
            let Some(IRNode::Slide(slide)) = store.get_mut(slide_id) else {
                continue;
            };
            for anim in &mut slide.animations {
                if let Some(target) = anim.target.as_ref() {
                    if let Some(id) = utils::resolve_media_asset(&media_by_path, target) {
                        anim.media_asset = Some(id);
                    }
                }
            }
        }
    }
}
