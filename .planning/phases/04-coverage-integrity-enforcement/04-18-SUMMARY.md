---
phase: 04-coverage-integrity-enforcement
plan: "18"
subsystem: testing
tags: [coverage, llvm-cov, parser, docir-diff, docir-core, docir-rules]
requires:
  - phase: 04-17
    provides: canonical 78.26% baseline and refreshed residual ranking
provides:
  - behavior-first branch tests for parser/rules utility hotspots
  - additional diff summary fallback/signature coverage
  - canonical 04-18 fail-under evidence and residual ranking
affects: [CC-04, TEST-01, TEST-02, phase-04-closure]
completed: 2026-03-01
---

# Phase 04 Plan 18 Summary

Wave 04-18 expanded behavior-first hotspot tests across diff/core/parser/rules, increasing canonical coverage from `78.26%` to `78.75%` (`+0.49` points), with fail-under remaining `EXIT:1`.

## Delivered

- Added branch-focused tests in:
  - `crates/docir-diff/src/summary.rs`
  - `crates/docir-core/src/visitor/visitors.rs`
  - `crates/docir-parser/src/rtf/objects.rs`
  - `crates/docir-parser/src/zip_handler.rs`
  - `crates/docir-rules/src/profile.rs`
- Captured explicit environment blocker for `docir-python` coverage in this runner (Python symbol link failures in unit-test binaries).
- Published canonical 04-18 measurement and refreshed missed-line ranking.

## Verification Commands

- `cargo test -p docir-diff -p docir-core -p docir-parser -p docir-rules`
- `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`

## Outcome

- Canonical total: `78.75%`
- Canonical fail-under status: `EXIT:1`
- Remaining threshold gap: `16.25` points
