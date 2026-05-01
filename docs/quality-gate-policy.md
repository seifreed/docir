# Quality Gate Non-Bypass Policy

## Purpose

This repository enforces one canonical quality gate surface to keep acceptance criteria deterministic and non-bypassable.

## Allowed Entrypoint

- Allowed entrypoint: `./scripts/quality_gate.sh`
- This is the only accepted command for final quality acceptance in local workflows and CI workflows.

## Additional mandatory controls (non-authoritative)

- Baseline and architecture measurements for Phase 1 are executed separately with:
  - `./scripts/quality_phase1_baseline.sh`
- The baseline output is mandatory for planning and audit traceability, but does not replace the canonical gate.
- Phase 2 operational control is executed through the canonical gate and includes:
  - `./scripts/quality_no_unwrap_expect_in_production.sh`
  - `./scripts/quality_no_wildcard_super_in_production.sh` (modo `working`: no nuevos `use super::*` en diff)
  - transición habilitada a inventario estricto con `QUALITY_NO_WILDCARD_INVENTORY_FAIL=1`
- Phase 3 boundary control is enforced by:
  - `./scripts/quality_layer_policy.sh`
  - `./scripts/quality_presentation_boundary.sh`
  - `./scripts/quality_dependency_cycles.sh`
- Phase 4 API hygiene control is enforced by:
  - `./scripts/quality_api_hygiene.sh` (CC-12..CC-14 + lint strictness for dead code/imports)

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

## Coverage Integrity Policy

- Coverage metrics are accepted only when tests validate behavior outcomes (parser nodes, diagnostics, security indicators, and exported coverage content contracts).
- Synthetic or tautological tests that only execute lines without asserting meaningful outputs are forbidden as acceptance evidence.
- Fixture-driven assertions over real sample documents are required for coverage-oriented changes.
- Direct coverage commands are diagnostic-only; canonical gate output remains the sole acceptance authority.
- Coverage threshold for `coverage_check` is sourced from `scripts/quality_coverage_threshold.txt` and can be temporarily overridden with `QUALITY_GATE_COVERAGE_THRESHOLD=<value>`.

## API Hygiene and Complexity Checks

- `scripts/quality_api_hygiene.sh` executes before `fmt/clippy/test/coverage`.
- Dead code/import strictness: enforced via strict linting in hygiene stage (`dead_code`, `unused_imports`).
- `CC-12`: enforced by scanning public function declarations for missing `///`/doc comments.
- `CC-13`: enforced by scanning function complexity; defaults to threshold `CC13_COMPLEXITY_THRESHOLD=10`.
- `CC-14`: enforced by scanning file/function LOC thresholds in production scope.
- Failure output is emitted as `api_hygiene policy: FAIL` for deterministic gate traceability.
- Objective threshold policy:
  - hard-fail operativo en `>100 LOC` para funciones (CC-14);
  - objetivo de madurez para 10/10: converger a `>80 LOC` o excepción justificada.

## Production Robustness Metric Scope

- Canonical scope for robustez de producción en inventario y baseline:
  - `crates/docir-parser/src/**`
  - `crates/docir-app/src/**`
  - `crates/docir-diff/src/**`
  - `crates/docir-security/src/**`
- Exclusiones obligatorias:
  - `tests/**`, `*_tests.rs`, `test.rs`, `tests.rs`, `tests_*`, `test_*`
- Rationale:
  - este scope mide rutas operativas del pipeline y servicios de aplicación;
  - evita mezclar deuda histórica de crates de interfaz/harness con riesgo operativo del parser stack.

## CI Evidence Registration

- CI ejecuta `./scripts/quality_gate.sh` como gate canónico.
- CI publica artefacto `quality-evidence` con:
  - `quality_phase1_baseline.log`
  - `quality_wildcard_inventory.log`
  - `residual_dashboard.md`

## Compliance Record

Current scan evidence is tracked in:

- `.planning/phases/01-canonical-gate-surface/01-NON_BYPASS_INVENTORY.md`
