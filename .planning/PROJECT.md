# docir Deterministic Quality Enforcement

## What This Is

This project hardens the existing `docir` Rust workspace into a strictly enforced Clean Code and Clean Architecture system with deterministic, non-bypassable quality gates. The repository already delivers document parsing and malware-analysis capabilities, and this effort adds a canonical enforcement surface that must pass locally and in CI. Completion is defined only by a single canonical gate pass where all requirements succeed simultaneously.

## Core Value

Quality and architecture compliance are deterministic, enforceable, and impossible to bypass through any alternate path.

## Requirements

### Validated

- ✓ Multi-crate Rust workspace with CLI and Python surfaces for document IR and security analysis — existing
- ✓ Parsing and analysis pipeline for Office/ODF/HWP/RTF documents — existing
- ✓ Existing test, lint, and build workflows via Cargo — existing
- ✓ Existing codebase mapping artifacts in `.planning/codebase/` for architectural context — existing
- ✓ Canonical quality-gate entrypoint exists only at `./scripts/quality_gate.sh` with deterministic exit behavior — Phase 1
- ✓ Canonical-only non-bypass policy is documented with inventory evidence — Phase 1
- ✓ Local, pre-commit, and CI workflows are routed through canonical gate with required check `quality-gate` on `main` — Phase 2

### Active

- [ ] Canonical gate enforces formatting, linting, tests, coverage threshold, and static policy checks in one run.
- [ ] Clean Code policy violations (unwrap/expect/panic/todo/unimplemented/dead code/unused imports/missing docs/complexity breaches) fail the gate.
- [ ] Clean Architecture dependency violations between domain/application/infrastructure/presentation fail the gate.
- [ ] CI required job executes the canonical gate and blocks merges when any check fails.
- [ ] Enforcement is iterative and demonstrable (at least one failing and one passing gate state during implementation).
- [ ] Documentation codifies policy, non-bypass rule, and definition of done.

### Out of Scope

- Adding new product features unrelated to quality-gate enforcement — preserves focus on architecture and maintainability controls.
- Lowering thresholds or introducing warning suppressions to achieve green builds — conflicts with deterministic enforcement intent.
- Parallel non-canonical scripts that can act as alternate quality entrypoints — violates single-gate contract.

## Context

`docir` is an existing Rust 2021 workspace with multiple crates (`docir-core`, `docir-parser`, `docir-security`, `docir-rules`, `docir-diff`, `docir-serialization`, `docir-app`, `docir-cli`, `docir-python`). The current architecture is modular and functional, but there are known quality and complexity concerns in parser-heavy areas and a need for consistent, repository-wide quality enforcement. This initiative is explicitly for this repository (not a reusable framework), and success requires deterministic gate behavior that is identical across local development and CI.

## Constraints

- **Gate Surface**: `./scripts/quality_gate.sh` only — all quality enforcement must route through this single entrypoint.
- **Determinism**: Single canonical execution outcome — repository is incomplete unless all checks pass simultaneously.
- **Quality Threshold**: Coverage must be at least 95% via `cargo llvm-cov` — no artificial test inflation accepted.
- **Policy Strictness**: No temporary ignores, no disabled warnings, no bypass scripts — enforcement must be durable.
- **Compatibility**: Keep behavior stable except for quality/architecture compliance changes — avoid unnecessary functional drift.

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Use `./scripts/quality_gate.sh` as sole quality gate surface | Eliminates ambiguity and bypass paths across local and CI workflows | ✓ Good |
| Treat gate failure as project-incomplete state | Enforces deterministic quality bar and prevents soft completion | — Pending |
| Enforce Clean Architecture layering as hard checks | Prevents architectural drift as workspace grows | — Pending |
| Scope enforcement to this repository only | User-selected scope prioritizes immediate reliability over framework generalization | — Pending |
| Enforce merge gating with required check `quality-gate` | Converts routing policy into platform-enforced merge control | ✓ Good |

---
*Last updated: 2026-02-28 after Phase 2*
