# Quality Gate Non-Bypass Policy

## Purpose

This repository enforces one canonical quality gate surface to keep acceptance criteria deterministic and non-bypassable.

## Allowed Entrypoint

- Allowed entrypoint: `./scripts/quality_gate.sh`
- This is the only accepted command for final quality acceptance in local workflows and CI workflows.

## Forbidden Alternate Entrypoints

The following are forbidden as accepted gate surfaces:

- Any script other than `./scripts/quality_gate.sh` presented as equivalent quality gate execution.
- Any workflow step that runs raw checks directly (`cargo fmt`, `cargo clippy`, `cargo test`, coverage, policy scans) as a substitute gate result.
- Any wrapper script, alias, Make target, or task runner command documented as an accepted bypass path.

## Bypass Pattern Examples

Examples of prohibited bypass patterns:

- "Run `cargo clippy` and `cargo test`; this is equivalent to the gate."
- "CI quality job runs direct Cargo commands instead of `./scripts/quality_gate.sh`."
- "Add `scripts/gate_fast.sh` and treat it as accepted quality approval."

## Enforcement Expectations

- New scripts, hooks, automation, and workflows that perform quality checks must route through the canonical gate command.
- Optional fast checks may exist for developer feedback, but they must be documented as non-authoritative and never equivalent to gate acceptance.
- Documentation and workflow definitions must not describe any alternate accepted gate path.

## Compliance Record

Current scan evidence is tracked in:

- `.planning/phases/01-canonical-gate-surface/01-NON_BYPASS_INVENTORY.md`
