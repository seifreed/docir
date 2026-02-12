//! Security policy constants and heuristics.

pub use docir_core::security::{
    analyze_vba_source, is_dangerous_xlm_function, VbaAnalysis, AUTO_EXEC_PROCEDURES,
    DANGEROUS_XLM_FUNCTIONS, SUSPICIOUS_VBA_CALLS,
};
