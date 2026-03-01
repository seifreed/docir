---
phase: 04-coverage-integrity-enforcement
plan: "14"
subsystem: testing
tags: [coverage, llvm-cov, cross-crate, cli, diff, core]
requires:
  - phase: 04-13
    provides: canonical 73.44% baseline and residual hotspot ranking
provides:
  - cross-crate hotspot tests for docir-diff/docir-core utilities
  - CLI summary/security behavior integration tests
  - canonical 04-14 evidence with wave-by-wave totals and blocker profile
affects: [CC-04, TEST-01, TEST-02, phase-04-closure]
completed: 2026-03-01
---

# Phase 04 Plan 14 Summary

Cross-crate behavior tests raised canonical workspace coverage from `73.44%` to `76.61%` (`+3.17` points), but canonical fail-under still exits `1`.

## Delivered

- Added broad behavior tests for `docir-diff` summary/index branches and `docir-core` visitor counters.
- Added CLI integration tests validating `summary` and `security` command behavior (human + JSON) on real fixtures.
- Published 04-14 canonical evidence with two-wave progression and full-workspace hotspot blocker ranking.

## Verification Commands

- `cargo test -p docir-diff -p docir-core`
- `cargo test -p docir-cli --test summary_security_commands`
- `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- `cargo llvm-cov report --json --summary-only --output-path target/llvm-cov-summary.json`

## Outcome

- Canonical total: `76.61%`
- Canonical fail-under status: `EXIT:1`
- Remaining threshold gap: `18.39` points

Phase 04 remains open until canonical fail-under reaches `EXIT:0` at `>=95%`.
