---
phase: 04-coverage-integrity-enforcement
plan: "15"
subsystem: testing
tags: [coverage, llvm-cov, cross-crate, docir-diff, docir-rules]
requires:
  - phase: 04-14
    provides: canonical 76.60% baseline and hotspot ranking
provides:
  - expanded branch tests for docir-diff summary primary/secondary paths
  - expanded synthetic behavior tests for docir-rules rules and burst thresholds
  - canonical 04-15 fail-under evidence with residual ranking
affects: [CC-04, TEST-01, TEST-02, phase-04-closure]
completed: 2026-03-01
---

# Phase 04 Plan 15 Summary

Cross-crate behavior tests increased canonical workspace coverage from `76.60%` to `77.67%` (`+1.07` points), but canonical fail-under still exits `1`.

## Delivered

- Added broad branch coverage tests in `docir-diff/src/summary.rs` for primary + secondary summary paths.
- Added synthetic behavior tests in `docir-rules/src/rules.rs` covering rule metadata, security findings, and burst thresholds.
- Published canonical 04-15 measurement and updated hotspot ranking.

## Verification Commands

- `cargo test -p docir-diff summary::tests -- --nocapture`
- `cargo test -p docir-rules -- --nocapture`
- `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`

## Outcome

- Canonical total: `77.67%`
- Canonical fail-under status: `EXIT:1`
- Remaining threshold gap: `17.33` points
