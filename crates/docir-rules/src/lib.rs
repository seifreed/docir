//! # docir-rules
//!
//! Rule engine and built-in rules for docir IR analysis.

mod engine;
mod profile;
mod rules;

pub use engine::{Finding, Rule, RuleCategory, RuleContext, RuleEngine, RuleReport, Severity};
pub use profile::{RuleProfile, RuleThresholds};

#[cfg(test)]
mod tests;
