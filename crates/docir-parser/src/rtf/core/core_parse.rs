//! RTF parsing support.

#[path = "core_parse_events.rs"]
mod core_parse_events;
pub(crate) use core_parse_events::*;
#[path = "core_parse_finalize.rs"]
mod core_parse_finalize;
pub(crate) use core_parse_finalize::*;

use crate::error::ParseError;
use docir_core::visitor::IrStore;

pub(crate) use super::cursor::RtfCursor;
use super::state::GroupKind;
pub(crate) use super::state::RtfParseContext;

pub(crate) fn parse_rtf(
    cursor: &mut RtfCursor<'_>,
    ctx: &mut RtfParseContext,
    store: &mut IrStore,
) -> Result<(), ParseError> {
    while !cursor.is_eof() {
        match cursor.next() {
            Some(b'{') => {
                ctx.push_group(GroupKind::Normal)?;
            }
            Some(b'}') => {
                flush_text(ctx, store, None)?;
                handle_group_end(ctx, store);
                ctx.pop_group()?;
            }
            Some(b'\\') => {
                if skip_unicode_fallback(cursor, ctx) {
                    continue;
                }
                parse_control(cursor, ctx, store)?;
            }
            Some(b'\r') | Some(b'\n') => {
                // ignore raw newlines
            }
            Some(_byte) if ctx.unicode_skip > 0 => {
                ctx.unicode_skip -= 1;
            }
            Some(byte) => match ctx.current_group_kind() {
                GroupKind::Normal | GroupKind::FieldResult | GroupKind::FieldInst => {
                    append_text_byte(ctx, byte);
                }
                GroupKind::Object | GroupKind::Picture => {
                    if ctx.current_group_kind() == GroupKind::Object
                        && ctx.object_text_target.is_some()
                    {
                        append_text_byte(ctx, byte);
                    } else if byte.is_ascii_hexdigit()
                        && let Some(obj) = ctx.object_stack.last_mut()
                    {
                        obj.data_hex_len += 1;
                        if ctx.max_object_hex_len > 0 && obj.data_hex_len > ctx.max_object_hex_len {
                            return Err(ParseError::ResourceLimit(format!(
                                "RTF objdata too large: {} hex chars (max: {})",
                                obj.data_hex_len, ctx.max_object_hex_len
                            )));
                        }
                    }
                }
                GroupKind::FontTable | GroupKind::ColorTable => {
                    append_text_byte(ctx, byte);
                }
                _ => {}
            },
            None => break,
        }
    }
    if ctx.group_stack.len() != 1 {
        return Err(ParseError::InvalidStructure(
            "RTF has unclosed groups".to_string(),
        ));
    }
    flush_text(ctx, store, None)?;
    finalize_table_if_open(ctx, store);
    finalize_paragraph(ctx, store);
    finalize_section(ctx, store);
    Ok(())
}

fn skip_unicode_fallback(cursor: &mut RtfCursor<'_>, ctx: &mut RtfParseContext) -> bool {
    if ctx.unicode_skip == 0 {
        return false;
    }

    let Some(next) = cursor.next() else {
        ctx.unicode_skip = 0;
        return true;
    };
    match next {
        b'\'' => {
            cursor.next();
            cursor.next();
        }
        b'\\' | b'{' | b'}' | b'*' => {}
        byte if byte.is_ascii_alphabetic() => {
            while cursor.peek().is_some_and(|byte| byte.is_ascii_alphabetic()) {
                cursor.next();
            }
            if cursor.peek() == Some(b'-') {
                cursor.next();
            }
            while cursor.peek().is_some_and(|byte| byte.is_ascii_digit()) {
                cursor.next();
            }
            if cursor.peek() == Some(b' ') {
                cursor.next();
            }
        }
        _ => {}
    }
    ctx.unicode_skip -= 1;
    true
}
