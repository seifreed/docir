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

        let direct_function = super::extract_formula_function(upper_text);
        for func in contains_dangerous_xlm(formula_text) {
            let args = direct_function
                .as_deref()
                .filter(|direct| *direct == func)
                .and_then(|_| super::parse_formula_args_text(formula_text));
            xlm.dangerous_functions.push(XlmFunction {
                name: func,
                arguments: args,
                cell_ref: cell_ref.to_string(),
            });
        }

        let _ = sheet_path;
    }
}
