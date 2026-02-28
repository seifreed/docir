# Project Research Summary

**Project:** docir Deterministic Quality Enforcement
**Domain:** Deterministic Rust quality-gate enforcement for a multi-crate malware-analysis workspace
**Researched:** 2026-02-28
**Confidence:** HIGH

## Executive Summary

This project is not a new runtime product; it is an enforcement system layered onto the existing `docir` Rust workspace so quality and architecture compliance become deterministic, non-bypassable, and merge-blocking. The expert pattern is clear across all research outputs: use one canonical orchestration surface (`./scripts/quality_gate.sh`), run a fixed check order (`fmt -> clippy -> test/check -> coverage -> policy`), and apply hard-fail semantics in both local and CI paths.

The recommended implementation approach is policy-first and contract-driven. Pin the Rust toolchain and gate tooling versions, codify clean-code and dependency-boundary rules as versioned policy files, and make CI required status depend only on the canonical script. Launch scope should focus on P1 table-stakes only: single gate surface, deterministic sequencing, hard clean-code and architecture checks, required CI integration, and 95% coverage enforcement.

The main risks are bypass reintroduction, semantic blind spots, and scaling friction. Specifically: parallel gate entrypoints can quietly invalidate enforcement, parser reliability issues can survive green gates if diagnostics are weak, and parser hotspot complexity can slow future policy rollout. Mitigation is phased: lock gate surface first, then strengthen failure semantics and policy coverage, then address memory/complexity/parity risks with targeted decomposition and contract tests.

## Key Findings

### Recommended Stack

The stack is mature and low-risk: pinned stable Rust toolchain via `rust-toolchain.toml`, canonical Cargo workspace commands, `rustfmt`, `clippy` with deny warnings, rustdoc lint enforcement, and `cargo-llvm-cov` with `llvm-tools-preview` for coverage gating. Determinism relies on strict version alignment and lockfile enforcement (`--locked`/`--frozen`), not custom build frameworks.

This means roadmap effort should prioritize orchestration and policy design over new technology adoption. The strongest recommendation is to avoid nightly dependencies, avoid mutating checks in CI (`clippy --fix`), and avoid any second quality script that can become an unofficial bypass path.

**Core technologies:**
- Rust toolchain pinning (`rustup` + `rust-toolchain.toml`): deterministic compiler/lint behavior — removes local/CI drift.
- Cargo workspace commands (`check`, `test`, `doc`): canonical validation surface — official, reproducible workspace execution.
- `rustfmt` + `clippy` (`-D warnings`): formatting and lint hard gates — stable enforcement of clean-code policy.
- Rustdoc linting (`RUSTDOCFLAGS="-D warnings" cargo doc --no-deps`): documentation-quality gate — prevents silent doc regressions.
- `cargo-llvm-cov` + `llvm-tools-preview`: coverage threshold enforcement — required for project’s 95% minimum.

### Expected Features

Feature research is explicit: P1 must deliver deterministic enforcement and non-bypass CI behavior; P2/P3 features are valuable but should not delay launch. Most differentiators (risk taxonomy, remediation hints, cross-format consistency gates) are best added after baseline gate stability is achieved.

**Must have (table stakes):**
- Single canonical gate entrypoint (`./scripts/quality_gate.sh`) — required for local/CI parity.
- Deterministic check ordering with fail-fast behavior — required for reproducible pass/fail.
- Hard clean-code policy enforcement — required to block known reliability hazards.
- Hard architecture dependency enforcement — required to stop layer drift.
- CI required merge-blocking integration — required for non-bypass governance.
- Coverage enforcement at >=95% — required quantitative quality floor.

**Should have (competitive):**
- Machine-readable gate artifacts (JSON/markdown) — improves auditability and trend visibility.
- Risk-domain violation taxonomy — turns lint failures into operational risk signals.
- Guided remediation hints — reduces time-to-fix without weakening strictness.
- Cross-format security extraction consistency checks — raises trust in malware signal parity.

**Defer (v2+):**
- Parser control-flow drift detectors — high complexity and false-positive risk early.
- Determinism audits for output identity contracts — defer until downstream reproducibility pressure increases.
- Historical dashboards/SLOs for gate health — defer until enough run history exists.

### Architecture Approach

Architecture should remain externalized and layered: Interface (developer/CI triggers) -> Orchestration (`scripts/quality_gate.sh`) -> Enforcement engines (format, lint, test, coverage, policy) -> Workspace crates. The key pattern is policy-as-code for both clean-code and dependency direction, enforced by one canonical runner. Internal scaling should use composable checks and caching behind the same script, never additional user-facing gate entrypoints.

**Major components:**
1. `scripts/quality_gate.sh` — canonical sequencing, deterministic exits, shared local/CI contract.
2. `scripts/checks/*` + `policy/*.toml` — clean-code and architecture policy enforcement as code.
3. CI required job (`ci/quality_gate.yml`) — merge-blocking execution of canonical gate only.
4. Existing `docir-*` crates — runtime functionality under enforcement, unchanged as architectural base.

### Critical Pitfalls

1. **Parallel gate surfaces** — remove alternate scripts and wire all paths to canonical gate.
2. **Silent parse failures under green gates** — require typed parse outcomes and diagnostics-backed tests.
3. **Memory amplification blind spots** — add large-input regression checks and single-buffer ownership patterns.
4. **God parser files** — decompose by responsibility and enforce size/complexity boundaries.
5. **Repeated XML loop drift** — centralize traversal helpers and add cross-format contract tests.

## Implications for Roadmap

Based on research, suggested phase structure:

### Phase 1: Canonical Gate Foundation
**Rationale:** All other work depends on one deterministic enforcement surface and merge-blocking CI contract.
**Delivers:** `scripts/quality_gate.sh` as sole entrypoint, fixed check order, fail-fast semantics, required CI wiring.
**Addresses:** Canonical gate entrypoint, deterministic ordering, required CI integration.
**Avoids:** Parallel gate surface bypass pitfall.

### Phase 2: Core Policy Enforcement (Clean Code + Architecture)
**Rationale:** Once the shell exists, enforce meaningful quality and dependency invariants as hard-fail rules.
**Delivers:** Clean-code checks (banned patterns/docs/dead/unused policy), architecture dependency matrix validation.
**Uses:** `clippy`, rustdoc lints, policy files, crate dependency graph checks.
**Implements:** Policy-as-code enforcement layer.
**Avoids:** Soft architecture checks, panic/unsafe pattern leakage into protected branches.

### Phase 3: Coverage and Determinism Hardening
**Rationale:** Coverage threshold and deterministic toolchain behavior are explicit project constraints and should be stabilized early.
**Delivers:** Pinned toolchain/components, locked/frozen cargo semantics in CI, >=95% coverage gate via `cargo-llvm-cov`.
**Addresses:** Coverage enforcement table-stake and reproducibility guarantees.
**Avoids:** Tooling/version drift and advisory coverage anti-pattern.

### Phase 4: Parser Reliability and Failure Semantics
**Rationale:** Green gates are insufficient if parser failures are swallowed; semantic reliability must be made observable.
**Delivers:** Typed parse outcomes, diagnostics requirements, corrupted-fixture regression suite.
**Addresses:** Reliability risks identified in pitfalls and concerns.
**Avoids:** Silent parse-failure pitfall.

### Phase 5: Scalability and Structural Stabilization
**Rationale:** After baseline enforcement, reduce hotspot fragility and scale limits in parser-heavy modules.
**Delivers:** Memory regression checks for large fixtures, decomposition of hotspot modules, shared XML traversal primitives, parity contract tests, security extraction consolidation.
**Addresses:** Medium-term maintainability and cross-format consistency.
**Avoids:** God files, memory amplification, XML behavioral drift, fragmented security extraction.

### Phase Ordering Rationale

- Dependencies enforce this order: canonical entrypoint first, then policy checks, then threshold hardening, then domain-specific reliability/scaling controls.
- Architecture findings favor keeping enforcement external to runtime crates while tightening contracts through policy-as-code.
- Pitfalls map naturally from governance failures (bypass) to semantic failures (silent parse issues) to scaling failures (memory/complexity drift).

### Research Flags

Phases likely needing deeper research during planning:
- **Phase 4:** Parse failure semantics need repository-specific fixture and diagnostics contract design.
- **Phase 5:** Shared XML primitives and security extraction consolidation are high-complexity, high-coupling changes.

Phases with standard patterns (skip research-phase):
- **Phase 1:** Canonical script + required CI wiring is well-established and already documented.
- **Phase 2:** Clean-code and architecture policy enforcement patterns are standard for Rust workspaces.
- **Phase 3:** Toolchain pinning and coverage gating have mature official guidance.

## Confidence Assessment

| Area | Confidence | Notes |
|------|------------|-------|
| Stack | HIGH | Based on official Rust/Cargo/Clippy/rustdoc coverage documentation and explicit project constraints. |
| Features | HIGH | Strong alignment between project requirements and quality-gate best practices; prioritization is clear. |
| Architecture | HIGH | Architecture pattern is straightforward and matches current workspace boundaries and constraints. |
| Pitfalls | HIGH | Risks are concrete, codebase-specific, and mapped to prevention phases with verification criteria. |

**Overall confidence:** HIGH

### Gaps to Address

- Coverage scope/exclusion policy details for the 95% bar need explicit lock-down during planning to prevent threshold disputes.
- Exact architecture rule matrix (allowed crate dependency edges) needs final authoritative definition before strict enforcement.
- Parser reliability acceptance criteria (what counts as `Missing` vs `Failed`) need formalization for deterministic diagnostics behavior.

## Sources

### Primary (HIGH confidence)
- `.planning/research/STACK.md` — stack and tooling recommendations, versioning and determinism practices.
- `.planning/research/FEATURES.md` — table stakes, differentiators, anti-features, and prioritization.
- `.planning/research/ARCHITECTURE.md` — canonical gate architecture, component boundaries, phase-aligned build order.
- `.planning/research/PITFALLS.md` — critical failure modes, warnings, and mitigation mapping.
- `.planning/PROJECT.md` — binding project constraints and scope.
- Rust toolchain docs: https://rust-lang.github.io/rustup/concepts/toolchains.html
- Cargo/Clippy/rustdoc/coverage docs referenced in STACK research.

### Secondary (MEDIUM confidence)
- `.planning/codebase/ARCHITECTURE.md` — current boundary context informing enforcement placement.
- `.planning/codebase/CONCERNS.md` — hotspot and reliability concerns informing phase risk prioritization.
- `.planning/codebase/STRUCTURE.md` — workspace organization context.

### Tertiary (LOW confidence)
- None.

---
*Research completed: 2026-02-28*
*Ready for roadmap: yes*
