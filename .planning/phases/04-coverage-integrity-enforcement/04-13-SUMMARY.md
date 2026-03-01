---
phase: 04-coverage-integrity-enforcement
plan: "13"
subsystem: testing
tags: [coverage, llvm-cov, parser, odf, ooxml]
requires:
  - phase: 04-12
    provides: canonical 72.86% baseline and residual hotspot ranking
provides:
  - behavior-first residual tests for inline, worksheet, helpers, and spreadsheet hotspots
  - canonical 04-13 coverage evidence with fail-under gate truth and module deltas
  - deterministic residual handoff ranking for follow-on closure
affects: [CC-04, TEST-01, TEST-02, phase-04-closure]
tech-stack:
  added: []
  patterns: [behavior-first branch assertions, canonical fail-under truth capture]
key-files:
  created:
    - .planning/phases/04-coverage-integrity-enforcement/04-13-COVERAGE.md
  modified:
    - crates/docir-parser/src/ooxml/docx/document/inline.rs
    - crates/docir-parser/src/ooxml/xlsx/worksheet.rs
    - crates/docir-parser/src/odf/helpers.rs
    - crates/docir-parser/src/odf/spreadsheet.rs
key-decisions:
  - "Kept CC-04 acceptance anchored to canonical cargo llvm-cov --fail-under-lines 95 exit semantics."
  - "Kept 04-13 strictly chained to the 04-12 residual ranking after plan-check blocker correction."
patterns-established:
  - "Use module-local behavior tests to close parser fallback and relation-resolution branches without broadening scope."
  - "Publish residual ranking every cycle from canonical fail-under output to preserve deterministic gap closure."
requirements-completed: [CC-04, TEST-01, TEST-02]
duration: 18m
completed: 2026-03-01
---

# Phase 04 Plan 13: Coverage Integrity Enforcement Summary

**Behavior-first residual hotspot tests increased canonical workspace coverage from 72.86% to 73.44%, with major worksheet branch closure, while canonical 95% enforcement remains blocking.**

## Accomplishments

- Added DOCX inline hyperlink-anchor behavior tests for relationship-backed and anchor-only targets.
- Added XLSX worksheet tests covering external links/connections ingestion and pivot-cache records relation parsing.
- Added ODF helper/spreadsheet tests for row parsing, truncated formatting behavior, unnamed validation handling, and fast-mode spreadsheet boundary behavior.
- Re-ran canonical coverage commands and published 04-13 evidence with fail-under truth, module snapshots, and residual handoff ranking.

## Verification Commands

- `cargo test -p docir-parser ooxml::docx::document::inline --all-features`
- `cargo test -p docir-parser ooxml::xlsx::worksheet --all-features`
- `cargo test -p docir-parser odf::helpers --all-features`
- `cargo test -p docir-parser odf::spreadsheet --all-features`
- `bash scripts/tests/quality_gate_coverage_commands.sh`
- `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- `cargo llvm-cov --workspace --all-features --summary-only`

## Coverage Outcome

- Baseline: `72.86%` (04-12 canonical fail-under)
- Current: `73.44%` (04-13 canonical fail-under)
- Delta: `+0.58` percentage points
- Canonical status: `FAIL` (`EXIT:1`, still `<95.00%`)

## Next Gap Handoff

Residual ranking from 04-13 targeted set:

1. `inline.rs` - `369` missed lines
2. `helpers.rs` - `235` missed lines
3. `spreadsheet.rs` - `135` missed lines
4. `worksheet.rs` - `115` missed lines

Phase 04 closure remains quantitatively blocked until canonical fail-under exits `0`.
