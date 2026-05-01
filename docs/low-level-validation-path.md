# Low-Level Validation Path

This document defines the minimum stable validation route for the newer low-level commands when the full Cargo test workflow is affected by intermittent toolchain hangs in this environment.

## Scope

Applies to:

- `inspect-sheet-records`
- `inspect-slide-records`
- `extract-flash`

And to their parser-layer readers:

- `xls_records`
- `ppt_records`

## Problem Pattern

The environment intermittently hangs during `rustc` link or crate-test binary startup for some workspace targets.

Observed behavior:

- `cargo test` reaches:
  - `Finished test profile ...`
  - `Running unittests src/lib.rs (...)`
- then the test binary may remain alive without meaningful CPU activity or additional output

This does not consistently correlate with a concrete failing assertion.

## Preferred Validation Order

Run commands strictly one at a time.

Do not run multiple `cargo test` or `cargo check` commands in parallel in this repository.

Recommended order:

1. `cargo test -p docir-parser xls_records --lib -- --nocapture`
2. `cargo test -p docir-parser ppt_records --lib -- --nocapture`
3. `cargo test -p docir-cli inspect_sheet_records -- --nocapture`
4. `cargo test -p docir-cli inspect_slide_records -- --nocapture`
5. `cargo test -p docir-cli extract_flash -- --nocapture`

Current known state:

- steps 1 and 2 have completed successfully
- the instability currently concentrates on steps 3-5, where a CLI target recompiles `docir-parser`

If parser crate tests hang after the test binary starts:

- record the hang as environment/toolchain instability
- continue with formatter/text/JSON command-level checks where possible

## Minimum Acceptance When Environment Is Unstable

If the full target test binary cannot complete reliably, the minimum acceptable evidence is:

- `cargo fmt`
- targeted code review of new command wiring
- parser-layer unit tests present and syntactically integrated
- CLI E2E tests present for JSON/text output
- roadmap state remains `partial`, not `partial validated`

## When To Upgrade Coverage State

Only move XLS/PPT/SWF low-level phases from:

- `partial`

to:

- `partial validated`

when the targeted sequence above completes successfully in a clean environment.
