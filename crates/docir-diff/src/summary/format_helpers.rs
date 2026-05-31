use docir_core::ir::{IRNode, Paragraph, ShapeText};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;
use sha2::Digest;

pub(crate) fn paragraph_text(para: &Paragraph, store: &IrStore) -> String {
    let mut out = String::new();
    for run_id in &para.runs {
        if let Some(node) = store.get(*run_id) {
            match node {
                IRNode::Run(run) => out.push_str(&run.text),
                IRNode::Hyperlink(link) => out.push_str(&runs_text(&link.runs, store)),
                _ => {}
            }
        }
    }
    out
}

pub(crate) fn runs_text(run_ids: &[NodeId], store: &IrStore) -> String {
    let mut out = String::new();
    for run_id in run_ids {
        if let Some(IRNode::Run(run)) = store.get(*run_id) {
            out.push_str(&run.text);
        }
    }
    out
}

pub(crate) fn shape_text(text: &ShapeText) -> String {
    let mut out = String::new();
    for (p_idx, para) in text.paragraphs.iter().enumerate() {
        if p_idx > 0 {
            out.push('\n');
        }
        for run in &para.runs {
            out.push_str(&run.text);
        }
    }
    out
}

pub(crate) fn opt_str(value: &Option<String>) -> String {
    value.clone().unwrap_or_else(|| "-".to_string())
}

pub(crate) fn opt_bool(value: Option<bool>) -> String {
    value
        .map(|v| if v { "true" } else { "false" }.to_string())
        .unwrap_or_else(|| "-".to_string())
}

pub(crate) fn opt_u32(value: Option<u32>) -> String {
    value
        .map(|v| v.to_string())
        .unwrap_or_else(|| "-".to_string())
}

pub(crate) fn abbreviate(value: &str, max: usize) -> String {
    if value.chars().count() <= max {
        return value.to_string();
    }
    let mut out = String::new();
    for (i, ch) in value.chars().enumerate() {
        if i + 1 > max {
            break;
        }
        out.push(ch);
    }
    out.push_str("...");
    out
}

pub(crate) fn format_float(value: f64) -> String {
    if (value.fract() - 0.0).abs() < f64::EPSILON {
        format!("{value:.0}")
    } else {
        format!("{value:.6}")
    }
}

pub(crate) fn short_hash(input: &str) -> String {
    let mut hasher = sha2::Sha256::new();
    hasher.update(input.as_bytes());
    let hash = hasher.finalize();
    to_hex(&hash[..8])
}

fn to_hex(data: &[u8]) -> String {
    data.iter().map(|b| format!("{b:02x}")).collect()
}
