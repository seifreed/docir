use crate::error::ParseError;
use crate::rtf::objects::{ObjectContext, ObjectTextTarget};
use docir_core::ir::{
    Border, BorderStyle, Indentation, LineSpacingRule, Paragraph, ParagraphBorders, Spacing,
    StyleSet, StyleType, Table, TableCell, TableCellProperties, TableRow, TextAlignment,
};
use docir_core::types::NodeId;
use encoding_rs::Encoding;
use std::collections::HashMap;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum GroupKind {
    Normal,
    FontTable,
    ColorTable,
    Stylesheet,
    StylesheetEntry,
    Info,
    Skip,
    Field,
    FieldInst,
    FieldResult,
    Picture,
    Object,
    ListTable,
    List,
    ListLevel,
    ListOverrideTable,
    ListOverride,
}

#[derive(Debug, Clone, Default)]
pub(super) struct RtfStyleState {
    pub(super) style_id: Option<String>,
    pub(super) bold: Option<bool>,
    pub(super) italic: Option<bool>,
    pub(super) underline: Option<bool>,
    pub(super) strike: Option<bool>,
    pub(super) font_size: Option<u32>,
    pub(super) font_index: Option<u32>,
    pub(super) color_index: Option<usize>,
    pub(super) highlight_index: Option<usize>,
    pub(super) vertical: Option<docir_core::ir::VerticalTextAlignment>,
}

#[derive(Debug, Clone)]
pub(super) struct GroupState {
    pub(super) kind: GroupKind,
    pub(super) style: RtfStyleState,
}

#[derive(Debug, Default)]
pub(super) struct FontTable {
    pub(super) fonts: HashMap<u32, String>,
}

#[derive(Debug, Default)]
pub(super) struct ColorTable {
    pub(super) colors: Vec<Option<String>>, // index 0 is default
}

#[derive(Debug, Default)]
pub(super) struct FieldContext {
    pub(super) instruction: String,
    pub(super) runs: Vec<NodeId>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum BorderTarget {
    Top,
    Bottom,
    Left,
    Right,
    InsideH,
    InsideV,
}

#[derive(Debug, Default)]
pub(super) struct StyleEntryContext {
    pub(super) style_id: Option<String>,
    pub(super) style_type: Option<StyleType>,
    pub(super) name_buf: String,
}

#[derive(Debug)]
pub(crate) struct RtfParseContext {
    pub(crate) sections: Vec<NodeId>,
    pub(super) current_section: Option<NodeId>,
    pub(super) current_paragraph: Option<Paragraph>,
    pub(super) current_table: Option<Table>,
    pub(super) current_row: Option<TableRow>,
    pub(super) current_cell: Option<TableCell>,
    pub(super) font_table: FontTable,
    pub(super) color_table: ColorTable,
    pub(super) encoding: &'static Encoding,
    pub(super) group_stack: Vec<GroupState>,
    pub(super) field_stack: Vec<FieldContext>,
    pub(super) object_stack: Vec<ObjectContext>,
    pub(crate) style_set: Option<StyleSet>,
    pub(super) current_style: Option<StyleEntryContext>,
    pub(super) list_overrides: HashMap<i32, i32>,
    pub(super) list_levels: HashMap<i32, u32>,
    pub(super) list_level_formats: HashMap<(i32, u32), String>,
    pub(super) current_list_id: Option<i32>,
    pub(super) current_list_level: u32,
    pub(super) current_list_override: Option<i32>,
    pub(super) current_list_override_list_id: Option<i32>,
    pub(super) pending_list_override: Option<i32>,
    pub(super) pending_list_level: Option<u32>,
    pub(super) pending_para_style: Option<String>,
    pub(super) pending_alignment: Option<TextAlignment>,
    pub(super) pending_indent: Indentation,
    pub(super) pending_spacing: Spacing,
    pub(super) pending_line_rule: Option<LineSpacingRule>,
    pub(super) pending_para_border_target: Option<BorderTarget>,
    pub(super) pending_para_border: Border,
    pub(super) pending_para_borders: ParagraphBorders,
    pub(super) row_cellx: Vec<i32>,
    pub(super) current_cell_index: usize,
    pub(super) pending_cell_props: Option<TableCellProperties>,
    pub(super) pending_border_target: Option<BorderTarget>,
    pub(super) pending_border: Border,
    pub(super) object_text_target: Option<ObjectTextTarget>,
    pub(crate) media_assets: Vec<NodeId>,
    pub(super) external_refs: Vec<NodeId>,
    pub(super) ole_objects: Vec<NodeId>,
    pub(super) current_text: String,
    pub(super) current_props: RtfStyleState,
    pub(super) max_group_depth: usize,
    pub(super) max_object_hex_len: usize,
}

impl RtfParseContext {
    pub(crate) fn new(max_group_depth: usize, max_object_hex_len: usize) -> Self {
        let mut ctx = Self {
            sections: Vec::new(),
            current_section: None,
            current_paragraph: None,
            current_table: None,
            current_row: None,
            current_cell: None,
            font_table: FontTable::default(),
            color_table: ColorTable::default(),
            encoding: encoding_rs::WINDOWS_1252,
            group_stack: Vec::new(),
            field_stack: Vec::new(),
            object_stack: Vec::new(),
            style_set: None,
            current_style: None,
            list_overrides: HashMap::new(),
            list_levels: HashMap::new(),
            list_level_formats: HashMap::new(),
            current_list_id: None,
            current_list_level: 0,
            current_list_override: None,
            current_list_override_list_id: None,
            pending_list_override: None,
            pending_list_level: None,
            pending_para_style: None,
            pending_alignment: None,
            pending_indent: Indentation::default(),
            pending_spacing: Spacing::default(),
            pending_line_rule: None,
            pending_para_border_target: None,
            pending_para_border: Border {
                style: BorderStyle::None,
                width: None,
                color: None,
            },
            pending_para_borders: ParagraphBorders::default(),
            row_cellx: Vec::new(),
            current_cell_index: 0,
            pending_cell_props: None,
            pending_border_target: None,
            pending_border: Border {
                style: BorderStyle::None,
                width: None,
                color: None,
            },
            object_text_target: None,
            media_assets: Vec::new(),
            external_refs: Vec::new(),
            ole_objects: Vec::new(),
            current_text: String::new(),
            current_props: RtfStyleState::default(),
            max_group_depth,
            max_object_hex_len,
        };
        ctx.group_stack.push(GroupState {
            kind: GroupKind::Normal,
            style: RtfStyleState::default(),
        });
        ctx
    }

    pub(super) fn current_group_kind(&self) -> GroupKind {
        if let Some(group) = self.group_stack.last() {
            group.kind
        } else {
            GroupKind::Normal
        }
    }

    pub(super) fn push_group(&mut self, kind: GroupKind) -> Result<(), ParseError> {
        if self.max_group_depth > 0 && self.group_stack.len() >= self.max_group_depth {
            return Err(ParseError::ResourceLimit(format!(
                "RTF max group depth exceeded: {}",
                self.max_group_depth
            )));
        }
        let style = self.current_props.clone();
        self.group_stack.push(GroupState { kind, style });
        Ok(())
    }

    pub(super) fn pop_group(&mut self) {
        if let Some(group) = self.group_stack.pop() {
            self.current_props = group.style;
        }
    }
}
