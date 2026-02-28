# Quality Gate Non-Bypass Policy

## Purpose

This repository enforces one canonical quality gate surface to keep acceptance criteria deterministic and non-bypassable.

## Allowed Entrypoint

- Allowed entrypoint: `./scripts/quality_gate.sh`
- This is the only accepted command for final quality acceptance in local workflows and CI workflows.

## Forbidden Alternate Entrypoints

The following are forbidden as accepted gate surfaces:

- Any script other than `./scripts/quality_gate.sh` presented as accepted quality gate execution.
- Any workflow step that runs raw checks directly (`cargo fmt`, `cargo clippy`, `cargo test`, coverage, policy scans) as a substitute gate result.
- Any wrapper script, alias, Make target, or task runner command documented as an accepted bypass path.

## Bypass Pattern Examples

Examples of prohibited bypass patterns:

- "Run `cargo clippy` and `cargo test`; this is accepted as the gate."
- "CI quality job runs direct Cargo commands instead of `./scripts/quality_gate.sh`."
- "Add `scripts/gate_fast.sh` and treat it as accepted quality approval."

## Enforcement Expectations

- New scripts, hooks, automation, and workflows that perform quality checks must route through the canonical gate command.
- Optional fast checks may exist for developer feedback, but they must be documented as non-authoritative and never accepted as gate approval.
- Documentation and workflow definitions must not describe any alternate accepted gate path.

## Warning and Suppression Policy

- Canonical linting posture is warning-strict and enforced by `cargo clippy --all-targets --all-features -- -D warnings`.
- Suppression mechanisms are forbidden for acceptance, including `-A` CLI flags, acceptance-time lint allows, and env-based bypass toggles.
- Raw Cargo commands remain diagnostic-only; only canonical gate output authorizes acceptance.

## Compliance Record

Current scan evidence is tracked in:

- `.planning/phases/01-canonical-gate-surface/01-NON_BYPASS_INVENTORY.md`
