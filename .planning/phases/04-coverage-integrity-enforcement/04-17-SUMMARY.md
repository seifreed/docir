---
phase: 04-coverage-integrity-enforcement
plan: "17"
subsystem: testing
tags: [coverage, llvm-cov, cross-crate, docir-diff]
requires:
  - phase: 04-16
    provides: canonical 78.12% baseline and refreshed residual ranking
provides:
  - expanded intrinsic-key/index tests for docir-diff
  - canonical 04-17 fail-under evidence with refreshed residual ranking
affects: [CC-04, TEST-01, TEST-02, phase-04-closure]
completed: 2026-03-01
---

# Phase 04 Plan 17 Summary

Cross-crate index branch tests increased canonical coverage from `78.12%` to `78.26%` (`+0.14` points), while fail-under remained `EXIT:1`.

## Delivered

- Added intrinsic-key branch tests in `docir-diff/src/index.rs` across word/presentation/security variants.
- Kept behavior assertions tied to deterministic keys and index stability outputs.
- Published canonical 04-17 measurement with refreshed hotspot ranking.

## Verification Commands

- `cargo test -p docir-diff index::tests -- --nocapture`
- `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`

## Outcome

- Canonical total: `78.26%`
- Canonical fail-under status: `EXIT:1`
- Remaining threshold gap: `16.74` points
