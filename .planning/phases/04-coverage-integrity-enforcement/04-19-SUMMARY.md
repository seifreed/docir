---
phase: 04-coverage-integrity-enforcement
plan: "19"
subsystem: testing
tags: [coverage, llvm-cov, parser, rtf]
requires:
  - phase: 04-18
    provides: canonical 78.75% baseline and residual hotspot ranking
provides:
  - dense rtf core helper/finalization behavior tests
  - canonical 04-19 fail-under evidence with refreshed hotspot ranking
affects: [CC-04, TEST-01, TEST-02, phase-04-closure]
completed: 2026-03-01
---

# Phase 04 Plan 19 Summary

Wave 04-19 targeted `rtf/core.rs` residual branches and increased canonical coverage from `78.75%` to `79.22%` (`+0.47` points), while fail-under remained `EXIT:1`.

## Delivered

- Added helper- and finalization-focused tests in `crates/docir-parser/src/rtf/core.rs` covering:
  - run-style and field control dispatch branches,
  - field/style/list finalization branches,
  - border application, cell width derivation, and run-property mapping.
- Re-ran parser-focused tests and canonical workspace fail-under.
- Published 04-19 coverage evidence with updated hotspot ordering.

## Verification Commands

- `cargo test -p docir-parser rtf::core::tests -- --nocapture`
- `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`

## Outcome

- Canonical total: `79.22%`
- Canonical fail-under status: `EXIT:1`
- Remaining threshold gap: `15.78` points
