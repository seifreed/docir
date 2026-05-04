use crate::make_indicator;
use docir_core::ir::{DefinedName, IRNode, SheetState};
use docir_core::security::{ThreatIndicator, ThreatIndicatorType, ThreatLevel};
use docir_core::visitor::IrStore;
use std::collections::HashMap;

pub(super) fn build_xlm_indicators(
    store: &IrStore,
    security: &docir_core::security::SecurityInfo,
) -> Vec<ThreatIndicator> {
    let mut indicators = Vec::new();
    let mut sheet_locations = HashMap::new();

    for node in store.values() {
        if let IRNode::Worksheet(sheet) = node {
            if let Some(span) = sheet.span.as_ref() {
                sheet_locations.insert(sheet.name.to_ascii_uppercase(), span.file_path.clone());
            }
        }
    }

    if security.xlm_macros.is_empty() {
        return indicators;
    }

    for macro_entry in &security.xlm_macros {
        let location = sheet_locations
            .get(&macro_entry.sheet_name.to_ascii_uppercase())
            .cloned();

        if macro_entry.has_auto_open {
            indicators.push(make_indicator(
                ThreatIndicatorType::XlmMacro,
                ThreatLevel::Critical,
                "XLM Auto_Open macro detected".to_string(),
                location.clone(),
                None,
            ));
        }

        if macro_entry.sheet_state != SheetState::Visible && !macro_entry.macro_cells.is_empty() {
            indicators.push(make_indicator(
                ThreatIndicatorType::HiddenMacroSheet,
                ThreatLevel::High,
                format!("Hidden macro sheet: {}", macro_entry.sheet_name),
                location.clone(),
                None,
            ));
        }

        for func in &macro_entry.dangerous_functions {
            indicators.push(make_indicator(
                ThreatIndicatorType::XlmMacro,
                ThreatLevel::Critical,
                format!("XLM macro function {} at {}", func.name, func.cell_ref),
                location.clone(),
                None,
            ));
        }
    }

    indicators
}

pub(super) fn apply_xlm_defined_name_targets(
    store: &IrStore,
    security: &mut docir_core::security::SecurityInfo,
    indicators: &mut Vec<ThreatIndicator>,
) {
    let mut targets: Vec<Option<String>> = Vec::new();
    let mut location: Option<String> = None;

    for node in store.values() {
        if let IRNode::DefinedName(name) = node {
            if let Some(target) = auto_open_target_from_defined_name(name) {
                if location.is_none() {
                    location = name.span.as_ref().map(|s| s.file_path.clone());
                }
                targets.push(target);
            }
        }
    }

    if targets.is_empty() || security.xlm_macros.is_empty() {
        return;
    }

    let has_unresolved_target = targets.iter().any(|t| t.is_none());
    let mut any_marked = false;
    for target in targets.iter().flatten() {
        let target_upper = target.to_ascii_uppercase();
        for macro_entry in security.xlm_macros.iter_mut() {
            if macro_entry.sheet_name.to_ascii_uppercase() == target_upper {
                macro_entry.has_auto_open = true;
                any_marked = true;
            }
        }
    }

    if has_unresolved_target && !any_marked {
        // Mark only the first macro as auto_open when the target is unresolved,
        // rather than flagging all macros as a false-positive cascade.
        if let Some(first) = security.xlm_macros.first_mut() {
            first.has_auto_open = true;
        }
    }

    indicators.push(make_indicator(
        ThreatIndicatorType::XlmMacro,
        ThreatLevel::Critical,
        "XLM Auto_Open defined name detected".to_string(),
        location,
        None,
    ));
}

fn auto_open_target_from_defined_name(name: &DefinedName) -> Option<Option<String>> {
    let upper = name.name.to_ascii_uppercase();
    if upper == "_XLNM.AUTO_OPEN" || upper == "AUTO_OPEN" || upper == "AUTO.OPEN" {
        let value = name.value.trim();
        if let Some((sheet, _)) = value.split_once('!') {
            let cleaned = sheet.trim().trim_matches('\'').to_string();
            if !cleaned.is_empty() {
                return Some(Some(cleaned));
            }
        }
        return Some(None);
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::{DefinedName, IRNode, SheetKind, Worksheet};
    use docir_core::security::{ThreatIndicatorType, XlmFunction, XlmMacro, XlmMacroCell};
    use docir_core::types::{NodeId, SourceSpan};
    use docir_core::visitor::IrStore;

    fn macro_entry(sheet_name: &str) -> XlmMacro {
        XlmMacro {
            sheet_name: sheet_name.to_string(),
            sheet_state: SheetState::Hidden,
            dangerous_functions: vec![XlmFunction {
                name: "EXEC".to_string(),
                arguments: Some("cmd".to_string()),
                cell_ref: "A1".to_string(),
            }],
            macro_cells: vec![XlmMacroCell {
                cell_ref: "A1".to_string(),
                formula: r#"=EXEC("cmd")"#.to_string(),
            }],
            has_auto_open: false,
        }
    }

    #[test]
    fn build_xlm_indicators_emits_auto_open_hidden_and_function_signals() {
        let mut store = IrStore::new();
        let mut sheet = Worksheet::new("MacroSheet", 1);
        sheet.kind = SheetKind::MacroSheet;
        sheet.span = Some(SourceSpan::new("xl/worksheets/sheet1.xml"));
        store.insert(IRNode::Worksheet(sheet));

        let mut security = docir_core::security::SecurityInfo::new();
        let mut entry = macro_entry("MacroSheet");
        entry.has_auto_open = true;
        security.xlm_macros.push(entry);

        let indicators = build_xlm_indicators(&store, &security);
        assert_eq!(indicators.len(), 3);
        assert!(indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::XlmMacro
                && i.description.contains("Auto_Open")
                && i.location.as_deref() == Some("xl/worksheets/sheet1.xml")
        }));
        assert!(indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::HiddenMacroSheet
                && i.description.contains("Hidden macro sheet")
        }));
        assert!(indicators.iter().any(|i| {
            i.indicator_type == ThreatIndicatorType::XlmMacro
                && i.description.contains("EXEC")
                && i.description.contains("A1")
        }));
    }

    #[test]
    fn apply_xlm_defined_name_targets_marks_matching_sheet() {
        let mut store = IrStore::new();
        let defined_name = DefinedName {
            id: NodeId::new(),
            name: "_xlnm.auto_open".to_string(),
            value: "'MacroSheet'!$A$1".to_string(),
            local_sheet_id: None,
            hidden: false,
            comment: None,
            span: Some(SourceSpan::new("xl/workbook.xml")),
        };
        store.insert(IRNode::DefinedName(defined_name));

        let mut security = docir_core::security::SecurityInfo::new();
        security.xlm_macros.push(macro_entry("MacroSheet"));
        security.xlm_macros.push(macro_entry("OtherSheet"));
        let mut indicators = Vec::new();
        apply_xlm_defined_name_targets(&store, &mut security, &mut indicators);

        assert!(security.xlm_macros[0].has_auto_open);
        assert!(!security.xlm_macros[1].has_auto_open);
        assert_eq!(indicators.len(), 1);
        assert_eq!(indicators[0].indicator_type, ThreatIndicatorType::XlmMacro);
        assert_eq!(indicators[0].location.as_deref(), Some("xl/workbook.xml"));
    }

    #[test]
    fn apply_xlm_defined_name_targets_does_not_fallback_when_sheet_not_found() {
        let mut store = IrStore::new();
        let defined_name = DefinedName {
            id: NodeId::new(),
            name: "_xlnm.auto_open".to_string(),
            value: "'NonExistentSheet'!$A$1".to_string(),
            local_sheet_id: None,
            hidden: false,
            comment: None,
            span: None,
        };
        store.insert(IRNode::DefinedName(defined_name));

        let mut security = docir_core::security::SecurityInfo::new();
        security.xlm_macros.push(macro_entry("SheetA"));
        security.xlm_macros.push(macro_entry("SheetB"));
        let mut indicators = Vec::new();
        apply_xlm_defined_name_targets(&store, &mut security, &mut indicators);

        assert!(!security.xlm_macros[0].has_auto_open);
        assert!(!security.xlm_macros[1].has_auto_open);
    }

    #[test]
    fn apply_xlm_defined_name_targets_falls_back_to_mark_all_when_unresolved() {
        let mut store = IrStore::new();
        let defined_name = DefinedName {
            id: NodeId::new(),
            name: "AUTO.OPEN".to_string(),
            value: "A1".to_string(),
            local_sheet_id: None,
            hidden: false,
            comment: None,
            span: None,
        };
        store.insert(IRNode::DefinedName(defined_name));

        let mut security = docir_core::security::SecurityInfo::new();
        security.xlm_macros.push(macro_entry("SheetA"));
        security.xlm_macros.push(macro_entry("SheetB"));
        let mut indicators = Vec::new();
        apply_xlm_defined_name_targets(&store, &mut security, &mut indicators);

        // Only first macro marked when target is unresolved (not all macros)
        assert!(security.xlm_macros[0].has_auto_open);
        assert!(!security.xlm_macros[1].has_auto_open);
        assert_eq!(indicators.len(), 1);
    }
}
