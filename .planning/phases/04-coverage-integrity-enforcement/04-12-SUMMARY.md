---
phase: 04-coverage-integrity-enforcement
plan: "12"
subsystem: testing
tags: [coverage, llvm-cov, parser, odf, ooxml]
requires:
  - phase: 04-11
    provides: canonical 71.67% baseline and residual hotspot ranking
provides:
  - behavior-first residual tests for spreadsheet and inline hotspots
  - canonical 04-12 coverage evidence with fail-under gate truth and module deltas
  - deterministic residual handoff ranking for follow-on closure
affects: [CC-04, TEST-01, TEST-02, phase-04-closure]
tech-stack:
  added: []
  patterns: [parallel-path behavior tests, canonical fail-under truth capture]
key-files:
  created:
    - .planning/phases/04-coverage-integrity-enforcement/04-12-COVERAGE.md
  modified:
    - crates/docir-parser/src/odf/spreadsheet.rs
    - crates/docir-parser/src/ooxml/docx/document/inline.rs
key-decisions:
  - "Kept CC-04 acceptance anchored to canonical cargo llvm-cov --fail-under-lines 95 exit semantics."
  - "Focused 04-12 on high-yield residual branch families in spreadsheet parallel/pivot and inline run-property logic."
patterns-established:
  - "Use targeted helper-level tests inside hotspot modules to unlock large missed-line reductions without scope broadening."
  - "Record canonical totals and targeted module deltas each cycle for deterministic residual prioritization."
requirements-completed: [CC-04, TEST-01, TEST-02]
duration: 8m
completed: 2026-03-01
---

# Phase 04 Plan 12: Coverage Integrity Enforcement Summary

**Behavior-first hotspot expansion improved canonical workspace coverage from 71.67% to 72.86% with major missed-line reduction in `spreadsheet.rs`, while canonical 95% enforcement remains blocking.**

## Accomplishments

- Added new parallel-path spreadsheet tests covering pivot link/cache behavior, empty-sheet fallback, and malformed pivot XML error handling.
- Added DOCX inline run-properties/numbering malformed-input tests to cover parser error and style-mapping branches.
- Re-ran canonical coverage commands and published 04-12 evidence with fail-under truth, module snapshots, and residual handoff ranking.

## Verification Commands

- `cargo test -p docir-parser odf::spreadsheet --all-features`
- `cargo test -p docir-parser ooxml::docx::document::inline --all-features`
- `cargo test -p docir-parser ooxml::xlsx::worksheet --all-features`
- `cargo test -p docir-parser odf::helpers --all-features`
- `bash scripts/tests/quality_gate_coverage_commands.sh`
- `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- `cargo llvm-cov --workspace --all-features --summary-only`

## Coverage Outcome

- Baseline: `71.67%` (04-11 canonical fail-under)
- Current: `72.86%` (04-12 canonical fail-under)
- Delta: `+1.19` percentage points
- Canonical status: `FAIL` (`EXIT:1`, still `<95.00%`)

## Next Gap Handoff

Residual ranking from 04-12 targeted set:

1. `inline.rs` - `352` missed lines
2. `worksheet.rs` - `289` missed lines
3. `helpers.rs` - `234` missed lines
4. `spreadsheet.rs` - `134` missed lines

Phase 04 closure remains quantitatively blocked until canonical fail-under exits `0`.
