//! XLM macro tracking helpers for XLSX parsing.

use super::XlsxParser;
use crate::ooxml::xlsx::workbook::SheetInfo;
use docir_core::security::{XlmFunction, XlmMacro, XlmMacroCell};
use docir_security::contains_dangerous_xlm;

impl XlsxParser {
    pub(super) fn begin_macro_sheet(&mut self, sheet: &SheetInfo) {
        let xlm = XlmMacro {
            sheet_name: sheet.name.clone(),
            sheet_state: sheet.state,
            dangerous_functions: Vec::new(),
            macro_cells: Vec::new(),
            has_auto_open: false,
        };
        self.security_info.xlm_macros.push(xlm);
        self.current_xlm_index = Some(self.security_info.xlm_macros.len() - 1);
    }

    pub(super) fn finalize_auto_open_targets(&mut self, auto_open_targets: &[Option<String>]) {
        if auto_open_targets.is_empty() || self.security_info.xlm_macros.is_empty() {
            return;
        }

        let mut any_marked = false;
        for target in auto_open_targets.iter().flatten() {
            let target_upper = target.to_ascii_uppercase();
            for macro_entry in self.security_info.xlm_macros.iter_mut() {
                if macro_entry.sheet_name.to_ascii_uppercase() == target_upper {
                    macro_entry.has_auto_open = true;
                    any_marked = true;
                }
            }
        }
        if !any_marked {
            for macro_entry in self.security_info.xlm_macros.iter_mut() {
                macro_entry.has_auto_open = true;
            }
        }
    }

    pub(super) fn record_xlm_formula(
        &mut self,
        cell_ref: &str,
        formula_text: &str,
        upper_text: &str,
        sheet_path: &str,
    ) {
        if self.current_sheet_kind != Some(super::SheetKind::MacroSheet) {
            return;
        }

        let Some(idx) = self.current_xlm_index else {
            return;
        };

        let Some(xlm) = self.security_info.xlm_macros.get_mut(idx) else {
            return;
        };

        xlm.macro_cells.push(XlmMacroCell {
            cell_ref: cell_ref.to_string(),
            formula: formula_text.to_string(),
        });

        if upper_text.contains("AUTO_OPEN") || upper_text.contains("AUTO.OPEN") {
            xlm.has_auto_open = true;
        }

        for func in contains_dangerous_xlm(formula_text) {
            xlm.dangerous_functions.push(XlmFunction {
                arguments: parse_formula_args_for_function(formula_text, &func),
                name: func,
                cell_ref: cell_ref.to_string(),
            });
        }

        let _ = sheet_path;
    }
}

fn parse_formula_args_for_function(formula: &str, function: &str) -> Option<String> {
    let upper = formula.to_ascii_uppercase();
    let function = function.to_ascii_uppercase();
    let mut search_from = 0;

    while let Some(relative_start) = upper[search_from..].find(&function) {
        let start = search_from + relative_start;
        let end = start + function.len();
        let before_is_identifier =
            start > 0 && is_formula_identifier_byte(upper.as_bytes()[start - 1]);
        let after_is_identifier = upper
            .as_bytes()
            .get(end)
            .is_some_and(|byte| is_formula_identifier_byte(*byte));
        if before_is_identifier || after_is_identifier {
            search_from = end;
            continue;
        }

        let mut open = end;
        while upper
            .as_bytes()
            .get(open)
            .is_some_and(u8::is_ascii_whitespace)
        {
            open += 1;
        }
        if upper.as_bytes().get(open) != Some(&b'(') {
            search_from = end;
            continue;
        }

        let bytes = formula.as_bytes();
        let mut depth = 0usize;
        let mut in_string = false;
        let mut index = open;
        while index < bytes.len() {
            match bytes[index] {
                b'"' => {
                    if in_string && bytes.get(index + 1) == Some(&b'"') {
                        index += 2;
                        continue;
                    }
                    in_string = !in_string;
                }
                b'(' if !in_string => depth += 1,
                b')' if !in_string => {
                    depth -= 1;
                    if depth == 0 {
                        return Some(formula[open + 1..index].to_string());
                    }
                }
                _ => {}
            }
            index += 1;
        }
        return None;
    }

    None
}

fn is_formula_identifier_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.' | b'$')
}
