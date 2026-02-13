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

    let mut any_marked = false;
    for target in &targets {
        if let Some(target) = target {
            let target_upper = target.to_ascii_uppercase();
            for macro_entry in security.xlm_macros.iter_mut() {
                if macro_entry.sheet_name.to_ascii_uppercase() == target_upper {
                    macro_entry.has_auto_open = true;
                    any_marked = true;
                }
            }
        }
    }

    if !any_marked {
        for macro_entry in security.xlm_macros.iter_mut() {
            macro_entry.has_auto_open = true;
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
        let val = name.value.trim();
        if let Some((sheet, _)) = val.split_once('!') {
            let cleaned = sheet.trim().trim_matches('\'').to_string();
            if !cleaned.is_empty() {
                return Some(Some(cleaned));
            }
        }
        return Some(None);
    }
    None
}
