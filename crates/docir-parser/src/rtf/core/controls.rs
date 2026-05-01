#[path = "controls_table.rs"]
mod controls_table;
use super::pending_numbering;
use super::state::{GroupKind, RtfParseContext, StyleEntryContext};
use crate::error::ParseError;
use crate::rtf::objects::{ObjectContext, ObjectTextTarget};
use docir_core::ir::{MediaType, StyleSet, StyleType};

fn set_last_group_kind(ctx: &mut RtfParseContext, kind: GroupKind) {
    if let Some(group) = ctx.group_stack.last_mut() {
        group.kind = kind;
    }
}

fn set_stylesheet_entry(ctx: &mut RtfParseContext, style_id: String, style_type: StyleType) {
    set_last_group_kind(ctx, GroupKind::StylesheetEntry);
    ctx.current_style = Some(StyleEntryContext {
        style_id: Some(style_id),
        style_type: Some(style_type),
        name_buf: String::new(),
    });
}

fn apply_pending_numbering_to_current_paragraph(ctx: &mut RtfParseContext) {
    let numbering = pending_numbering(ctx);
    if let Some(para) = ctx.current_paragraph.as_mut() {
        if let Some(numbering) = numbering {
            para.properties.numbering = Some(numbering);
        }
    }
}

fn set_last_object_media_image(ctx: &mut RtfParseContext) {
    if let Some(obj) = ctx.object_stack.last_mut() {
        obj.media_type = Some(MediaType::Image);
    }
}

fn set_last_object_dimension(
    ctx: &mut RtfParseContext,
    param: Option<i32>,
    mut apply: impl FnMut(&mut ObjectContext, u32),
) {
    if let Some(value) = param {
        if let Some(obj) = ctx.object_stack.last_mut() {
            apply(obj, value.max(0) as u32);
        }
    }
}

pub(super) fn handle_group_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
) -> Result<bool, ParseError> {
    match word {
        "fonttbl" => {
            set_last_group_kind(ctx, GroupKind::FontTable);
            ctx.current_text.clear();
        }
        "colortbl" => {
            set_last_group_kind(ctx, GroupKind::ColorTable);
            if ctx.color_table.colors.is_empty() {
                ctx.color_table.colors.push(None);
            }
            ctx.current_text.clear();
        }
        "stylesheet" => {
            set_last_group_kind(ctx, GroupKind::Stylesheet);
            if ctx.style_set.is_none() {
                ctx.style_set = Some(StyleSet::new());
            }
        }
        "s" => {
            if let Some(id) = param {
                let style_id = format!("s{}", id.max(0));
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    set_stylesheet_entry(ctx, style_id, StyleType::Paragraph);
                } else {
                    ctx.pending_para_style = Some(style_id);
                    if let Some(para) = ctx.current_paragraph.as_mut() {
                        para.style_id = ctx.pending_para_style.clone();
                    }
                }
            }
        }
        "cs" => {
            if let Some(id) = param {
                let style_id = format!("cs{}", id.max(0));
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    set_stylesheet_entry(ctx, style_id, StyleType::Character);
                } else {
                    ctx.current_props.style_id = Some(style_id);
                }
            }
        }
        "ds" => {
            if let Some(id) = param {
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    set_stylesheet_entry(ctx, format!("ds{}", id.max(0)), StyleType::Other);
                }
            }
        }
        "ts" => {
            if let Some(id) = param {
                if ctx.current_group_kind() == GroupKind::Stylesheet {
                    set_stylesheet_entry(ctx, format!("ts{}", id.max(0)), StyleType::Table);
                }
            }
        }
        "info" => {
            set_last_group_kind(ctx, GroupKind::Info);
        }
        "listtable" => {
            set_last_group_kind(ctx, GroupKind::ListTable);
        }
        "listoverridetable" => {
            set_last_group_kind(ctx, GroupKind::ListOverrideTable);
        }
        "list" => {
            if ctx.current_group_kind() == GroupKind::ListTable {
                set_last_group_kind(ctx, GroupKind::List);
                ctx.current_list_id = None;
                ctx.current_list_level = 0;
            }
        }
        "listlevel" => {
            if ctx.current_group_kind() == GroupKind::List {
                ctx.current_list_level = ctx.current_list_level.saturating_add(1);
                set_last_group_kind(ctx, GroupKind::ListLevel);
            }
        }
        "levelnfc" => {
            if let (Some(list_id), Some(nfc)) = (ctx.current_list_id, param) {
                let level = ctx.current_list_level.saturating_sub(1);
                ctx.list_level_formats
                    .insert((list_id, level), format!("nfc:{}", nfc));
            }
        }
        "listid" => {
            if let Some(id) = param {
                if matches!(
                    ctx.current_group_kind(),
                    GroupKind::List | GroupKind::ListLevel | GroupKind::ListOverride
                ) {
                    ctx.current_list_id = Some(id);
                    if ctx.current_group_kind() == GroupKind::ListOverride {
                        ctx.current_list_override_list_id = Some(id);
                    }
                }
            }
        }
        "listoverride" => {
            if ctx.current_group_kind() == GroupKind::ListOverrideTable {
                set_last_group_kind(ctx, GroupKind::ListOverride);
                ctx.current_list_override = None;
                ctx.current_list_override_list_id = None;
            }
        }
        "ls" => {
            if let Some(id) = param {
                if ctx.current_group_kind() == GroupKind::ListOverride {
                    ctx.current_list_override = Some(id);
                } else {
                    ctx.pending_list_override = Some(id);
                    apply_pending_numbering_to_current_paragraph(ctx);
                }
            }
        }
        "ilvl" => {
            if let Some(level) = param {
                ctx.pending_list_level = Some(level.max(0) as u32);
                apply_pending_numbering_to_current_paragraph(ctx);
            }
        }
        _ => return Ok(false),
    }
    Ok(true)
}

pub(super) fn handle_object_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
) -> bool {
    match word {
        "pict" => {
            set_last_group_kind(ctx, GroupKind::Picture);
            ctx.object_stack.push(ObjectContext::default());
        }
        "pngblip" => {
            set_last_object_media_image(ctx);
        }
        "jpegblip" | "jpgblip" => {
            set_last_object_media_image(ctx);
        }
        "wmetafile" | "emfblip" | "wmetafile8" => {
            set_last_object_media_image(ctx);
        }
        "picw" => {
            set_last_object_dimension(ctx, param, |obj, value| {
                obj.pic_width = Some(value);
            });
        }
        "pich" => {
            set_last_object_dimension(ctx, param, |obj, value| {
                obj.pic_height = Some(value);
            });
        }
        "picwgoal" => {
            set_last_object_dimension(ctx, param, |obj, value| {
                obj.pic_width = Some(value);
            });
        }
        "pichgoal" => {
            set_last_object_dimension(ctx, param, |obj, value| {
                obj.pic_height = Some(value);
            });
        }
        "object" => {
            set_last_group_kind(ctx, GroupKind::Object);
            ctx.object_stack.push(ObjectContext::default());
        }
        "objclass" => {
            ctx.current_text.clear();
            ctx.object_text_target = Some(ObjectTextTarget::Class);
        }
        "objname" => {
            ctx.current_text.clear();
            ctx.object_text_target = Some(ObjectTextTarget::Name);
        }
        "objdata" => {
            set_last_group_kind(ctx, GroupKind::Object);
            ctx.current_text.clear();
            ctx.object_text_target = None;
        }
        _ => return false,
    }
    true
}

pub(super) fn handle_table_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
    store: &mut docir_core::visitor::IrStore,
) -> Result<bool, ParseError> {
    controls_table::handle_table_controls(word, param, ctx, store)
}

#[cfg(test)]
fn handle_table_cell_property_controls(
    word: &str,
    param: Option<i32>,
    ctx: &mut RtfParseContext,
) -> bool {
    controls_table::handle_table_cell_property_controls(word, param, ctx)
}

#[cfg(test)]
fn handle_table_border_controls(word: &str, param: Option<i32>, ctx: &mut RtfParseContext) -> bool {
    controls_table::handle_table_border_controls(word, param, ctx)
}

#[cfg(test)]
mod tests {
    use super::super::state::{BorderTarget, GroupKind, RtfParseContext};
    use super::{
        handle_group_controls, handle_object_controls, handle_table_border_controls,
        handle_table_cell_property_controls, handle_table_controls,
    };
    use crate::rtf::objects::ObjectTextTarget;
    use docir_core::ir::{BorderStyle, CellVerticalAlignment, MergeType, Paragraph};
    use docir_core::visitor::IrStore;

    fn ctx() -> RtfParseContext {
        RtfParseContext::new(128, 0)
    }

    #[test]
    fn handle_group_controls_updates_styles_and_list_state() {
        let mut ctx = ctx();

        assert!(handle_group_controls("fonttbl", None, &mut ctx).expect("fonttbl"));
        assert_eq!(ctx.current_group_kind(), GroupKind::FontTable);

        assert!(handle_group_controls("colortbl", None, &mut ctx).expect("colortbl"));
        assert_eq!(ctx.current_group_kind(), GroupKind::ColorTable);
        assert_eq!(ctx.color_table.colors.first(), Some(&None));

        assert!(handle_group_controls("stylesheet", None, &mut ctx).expect("stylesheet"));
        assert_eq!(ctx.current_group_kind(), GroupKind::Stylesheet);
        assert!(ctx.style_set.is_some());

        assert!(handle_group_controls("s", Some(3), &mut ctx).expect("s"));
        assert_eq!(
            ctx.current_style
                .as_ref()
                .and_then(|s| s.style_id.as_deref()),
            Some("s3")
        );

        assert!(handle_group_controls("cs", Some(4), &mut ctx).expect("cs"));
        assert_eq!(
            ctx.current_style
                .as_ref()
                .and_then(|s| s.style_id.as_deref()),
            Some("s3")
        );
        assert_eq!(ctx.current_props.style_id.as_deref(), Some("cs4"));

        assert!(handle_group_controls("listtable", None, &mut ctx).expect("listtable"));
        assert_eq!(ctx.current_group_kind(), GroupKind::ListTable);
        assert!(handle_group_controls("list", None, &mut ctx).expect("list"));
        assert_eq!(ctx.current_group_kind(), GroupKind::List);
        assert!(handle_group_controls("listid", Some(9), &mut ctx).expect("listid"));
        assert!(handle_group_controls("listlevel", None, &mut ctx).expect("listlevel"));
        assert!(handle_group_controls("levelnfc", Some(23), &mut ctx).expect("levelnfc"));
        assert_eq!(
            ctx.list_level_formats.get(&(9, 0)).map(String::as_str),
            Some("nfc:23")
        );

        assert!(handle_group_controls("listoverridetable", None, &mut ctx).expect("overrides"));
        assert!(handle_group_controls("listoverride", None, &mut ctx).expect("override"));
        assert!(handle_group_controls("listid", Some(9), &mut ctx).expect("override listid"));
        assert!(handle_group_controls("ls", Some(7), &mut ctx).expect("ls override"));
        assert_eq!(ctx.current_list_override, Some(7));

        // Outside override group, ls/ilvl update pending numbering fields.
        ctx.group_stack.clear();
        ctx.group_stack.push(crate::rtf::core::state::GroupState {
            kind: GroupKind::Normal,
            style: crate::rtf::core::state::RtfStyleState::default(),
        });
        ctx.current_paragraph = Some(Paragraph::new());
        assert!(handle_group_controls("ls", Some(5), &mut ctx).expect("ls pending"));
        assert!(handle_group_controls("ilvl", Some(2), &mut ctx).expect("ilvl"));
        assert_eq!(ctx.pending_list_override, Some(5));
        assert_eq!(ctx.pending_list_level, Some(2));
        assert!(!handle_group_controls("not-a-group-control", None, &mut ctx).expect("unknown"));
    }

    #[test]
    fn handle_object_controls_tracks_picture_and_object_metadata() {
        let mut ctx = ctx();

        assert!(handle_object_controls("pict", None, &mut ctx));
        assert_eq!(ctx.current_group_kind(), GroupKind::Picture);
        assert_eq!(ctx.object_stack.len(), 1);

        assert!(handle_object_controls("jpegblip", None, &mut ctx));
        assert_eq!(
            ctx.object_stack.last().and_then(|obj| obj.media_type),
            Some(docir_core::ir::MediaType::Image)
        );

        assert!(handle_object_controls("picw", Some(640), &mut ctx));
        assert!(handle_object_controls("pichgoal", Some(480), &mut ctx));
        assert_eq!(
            ctx.object_stack.last().and_then(|obj| obj.pic_width),
            Some(640)
        );
        assert_eq!(
            ctx.object_stack.last().and_then(|obj| obj.pic_height),
            Some(480)
        );

        assert!(handle_object_controls("object", None, &mut ctx));
        assert_eq!(ctx.current_group_kind(), GroupKind::Object);
        assert!(handle_object_controls("objclass", None, &mut ctx));
        assert_eq!(ctx.object_text_target, Some(ObjectTextTarget::Class));
        assert!(handle_object_controls("objname", None, &mut ctx));
        assert_eq!(ctx.object_text_target, Some(ObjectTextTarget::Name));
        assert!(handle_object_controls("objdata", None, &mut ctx));
        assert_eq!(ctx.object_text_target, None);

        assert!(!handle_object_controls(
            "unknown-object-control",
            None,
            &mut ctx
        ));
    }

    #[test]
    fn table_cell_and_border_controls_set_pending_properties() {
        let mut ctx = ctx();
        ctx.color_table.colors.push(Some("FF00AA".to_string()));

        assert!(handle_table_cell_property_controls(
            "clvmgf", None, &mut ctx
        ));
        assert!(handle_table_cell_property_controls(
            "clgridspan",
            Some(3),
            &mut ctx
        ));
        assert!(handle_table_cell_property_controls(
            "clcbpat",
            Some(0),
            &mut ctx
        ));
        assert!(handle_table_cell_property_controls(
            "clvertalc",
            None,
            &mut ctx
        ));
        assert!(!handle_table_cell_property_controls(
            "unknown-cell",
            None,
            &mut ctx
        ));

        let cell_props = ctx.pending_cell_props.as_ref().expect("cell props");
        assert_eq!(cell_props.vertical_merge, Some(MergeType::Restart));
        assert_eq!(cell_props.grid_span, Some(3));
        assert_eq!(cell_props.shading.as_deref(), Some("FF00AA"));
        assert_eq!(
            cell_props.vertical_align,
            Some(CellVerticalAlignment::Center)
        );

        assert!(handle_table_border_controls("clbrdrt", None, &mut ctx));
        assert!(handle_table_border_controls("brdrdb", None, &mut ctx));
        assert!(handle_table_border_controls("brdrw", Some(12), &mut ctx));
        assert!(handle_table_border_controls("brdrcf", Some(0), &mut ctx));
        assert!(handle_table_border_controls("brdrt", None, &mut ctx));
        assert!(!handle_table_border_controls(
            "unknown-border",
            None,
            &mut ctx
        ));

        assert_eq!(ctx.pending_border.style, BorderStyle::Double);
        assert_eq!(ctx.pending_border.width, Some(12));
        assert_eq!(ctx.pending_border.color.as_deref(), Some("FF00AA"));
        assert_eq!(ctx.pending_para_border_target, Some(BorderTarget::Top));
    }

    #[test]
    fn handle_group_controls_covers_info_and_stylesheet_variants() {
        let mut ctx = ctx();

        assert!(handle_group_controls("stylesheet", None, &mut ctx).expect("stylesheet"));
        assert!(handle_group_controls("ds", Some(2), &mut ctx).expect("ds"));
        assert_eq!(
            ctx.current_style
                .as_ref()
                .and_then(|s| s.style_id.as_deref()),
            Some("ds2")
        );
        assert!(handle_group_controls("stylesheet", None, &mut ctx).expect("stylesheet"));
        assert!(handle_group_controls("ts", Some(5), &mut ctx).expect("ts"));
        assert_eq!(
            ctx.current_style
                .as_ref()
                .and_then(|s| s.style_id.as_deref()),
            Some("ts5")
        );

        assert!(handle_group_controls("info", None, &mut ctx).expect("info"));
        assert_eq!(ctx.current_group_kind(), GroupKind::Info);
    }

    #[test]
    fn handle_table_and_object_controls_cover_remaining_variants() {
        let mut ctx = ctx();
        let mut store = IrStore::new();
        ctx.color_table.colors.push(Some("00AAFF".to_string()));

        assert!(handle_table_controls("trowd", None, &mut ctx, &mut store).expect("trowd"));
        assert!(ctx.current_table.is_some());
        assert!(ctx.current_row.is_some());
        assert!(ctx.current_cell.is_some());

        assert!(handle_table_controls("cellx", Some(1000), &mut ctx, &mut store).expect("cellx"));
        assert!(handle_table_controls("clvertalt", None, &mut ctx, &mut store).expect("clvertalt"));
        assert!(handle_table_controls("clvertalb", None, &mut ctx, &mut store).expect("clvertalb"));
        assert!(handle_table_controls("clvmrg", None, &mut ctx, &mut store).expect("clvmrg"));
        assert!(
            handle_table_controls("clgridspan", Some(0), &mut ctx, &mut store).expect("gridspan")
        );
        assert!(handle_table_controls("brdrtriple", None, &mut ctx, &mut store).expect("triple"));

        let cell_props = ctx.pending_cell_props.as_ref().expect("cell props");
        assert_eq!(
            cell_props.vertical_align,
            Some(CellVerticalAlignment::Bottom)
        );
        assert_eq!(cell_props.vertical_merge, Some(MergeType::Continue));
        assert_eq!(cell_props.grid_span, Some(1));
        assert_eq!(ctx.pending_border.style, BorderStyle::Triple);

        assert!(handle_table_controls("cell", None, &mut ctx, &mut store).expect("cell"));
        assert!(handle_table_controls("row", None, &mut ctx, &mut store).expect("row"));
        assert!(!handle_table_controls("unknown", None, &mut ctx, &mut store).expect("unknown"));

        assert!(handle_object_controls("pict", None, &mut ctx));
        assert!(handle_object_controls("pngblip", None, &mut ctx));
        assert!(handle_object_controls("wmetafile8", None, &mut ctx));
        assert!(handle_object_controls("picwgoal", Some(-1), &mut ctx));
        assert!(handle_object_controls("pich", Some(-2), &mut ctx));

        let obj = ctx.object_stack.last().expect("object context");
        assert_eq!(obj.media_type, Some(docir_core::ir::MediaType::Image));
        assert_eq!(obj.pic_width, Some(0));
        assert_eq!(obj.pic_height, Some(0));
    }
}
