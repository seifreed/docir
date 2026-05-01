use super::core_parse::{parse_control, parse_control_word_and_param};
use super::*;
use crate::error::ParseError;
use crate::rtf::core::state::GroupKind;
use docir_core::ir::IRNode;
use docir_core::visitor::IrStore;

#[test]
fn parse_control_word_and_param_handles_signed_numeric_parameter() {
    let mut cursor = RtfCursor::new(b"i-42 ");
    let first = cursor.next().expect("expected first control word byte");
    let (word, param) = parse_control_word_and_param(&mut cursor, first);
    assert_eq!(word, "i");
    assert_eq!(param, Some(-42));
    assert_eq!(cursor.peek(), None);
}

#[test]
fn parse_control_parses_quoted_hex_to_text() {
    let mut cursor = RtfCursor::new(b"'e4");
    let mut ctx = RtfParseContext::new(32, 0);
    let mut store = IrStore::new();

    parse_control(&mut cursor, &mut ctx, &mut store).expect("hex control should parse");

    assert_eq!(ctx.current_text, "ä");
}

#[test]
fn parse_control_updates_field_context_on_control_word() {
    let mut cursor = RtfCursor::new(b"field");
    let mut ctx = RtfParseContext::new(32, 0);
    let mut store = IrStore::new();

    parse_control(&mut cursor, &mut ctx, &mut store).expect("field control should parse");

    assert_eq!(
        ctx.group_stack.last().map(|g| g.kind),
        Some(GroupKind::Field)
    );
    assert_eq!(ctx.field_stack.len(), 1);
}

#[test]
fn parse_rtf_emits_style_runs_in_nested_groups() -> Result<(), ParseError> {
    let mut ctx = RtfParseContext::new(32, 0);
    let mut store = IrStore::new();
    let data = b"{\\rtf1{\\b Bold}{\\i1 italic}}";

    let mut cursor = RtfCursor::new(data);
    parse_rtf(&mut cursor, &mut ctx, &mut store)?;

    let run_styles: Vec<_> = store
        .iter()
        .filter_map(|(_, node)| match node {
            IRNode::Run(run) => {
                Some((run.text.clone(), run.properties.bold, run.properties.italic))
            }
            _ => None,
        })
        .collect();

    assert!(run_styles.len() >= 2);
    assert!(run_styles
        .iter()
        .any(|(_, bold, _)| bold == &Some(true) && bold.is_some()));
    assert!(run_styles
        .iter()
        .any(|(_, _, italic)| italic == &Some(true) && italic.is_some()));

    Ok(())
}
