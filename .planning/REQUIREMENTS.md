# Requirements: docir Deterministic Quality Enforcement

**Defined:** 2026-02-28
**Core Value:** Quality and architecture compliance are deterministic, enforceable, and impossible to bypass through any alternate path.

## v1 Requirements

Requirements for initial release. Each maps to roadmap phases.

### Canonical Gate Surface

- [ ] **GATE-01**: Repository provides exactly one canonical quality gate entrypoint at `./scripts/quality_gate.sh`.
- [ ] **GATE-02**: Canonical gate returns exit code `0` only when all checks pass and returns non-zero when any check fails.
- [ ] **GATE-03**: Local development quality workflow is documented and routed through the canonical gate only.
- [ ] **GATE-04**: Pre-commit quality workflow is documented and routed through the canonical gate only.
- [ ] **GATE-05**: CI required checks execute the canonical gate script directly.
- [ ] **GATE-06**: Repository contains no alternate or bypass quality scripts that can replace canonical gate execution.

### Clean Code Hard Gates

- [x] **CC-01**: Gate enforces `cargo fmt --all --check`.
- [x] **CC-02**: Gate enforces `cargo clippy --all-targets --all-features -- -D warnings`.
- [x] **CC-03**: Gate enforces `cargo test`.
- [x] **CC-04**: Gate enforces test coverage of at least 95% using `cargo llvm-cov`.
- [ ] **CC-05**: Gate fails if `unwrap()` appears in non-test code.
- [ ] **CC-06**: Gate fails if `expect()` appears in non-test code.
- [ ] **CC-07**: Gate fails if `panic!()` appears in production code.
- [ ] **CC-08**: Gate fails if `todo!()` appears in repository code.
- [ ] **CC-09**: Gate fails if `unimplemented!()` appears in repository code.
- [ ] **CC-10**: Gate fails when dead code is detected.
- [ ] **CC-11**: Gate fails when unused imports are detected.
- [ ] **CC-12**: Gate fails when any public function lacks documentation.
- [ ] **CC-13**: Gate fails when measurable cyclomatic complexity exceeds 10 for any function.

### Clean Architecture Hard Gates

- [ ] **ARCH-01**: Architecture policy defines layers `domain`, `application`, `infrastructure`, and `presentation` for workspace components.
- [ ] **ARCH-02**: Gate fails if domain layer depends on infrastructure layer.
- [ ] **ARCH-03**: Gate fails if domain layer depends on presentation layer.
- [ ] **ARCH-04**: Gate fails if application layer depends on presentation layer.
- [ ] **ARCH-05**: Gate fails if circular dependencies exist between crates.
- [ ] **ARCH-06**: Cross-layer boundaries are enforced through trait-based dependency inversion.
- [ ] **ARCH-07**: Gate fails if external-framework crates leak into domain layer.

### Testing and Integrity

- [x] **TEST-01**: Coverage target (>=95%) is measured in the canonical run and cannot be skipped in CI.
- [x] **TEST-02**: Tests used to satisfy gate requirements validate real behavior and are not trivial assertion-only inflation.
- [x] **TEST-03**: Warnings and lint checks remain fully enabled; warning suppression is not used to pass the gate.

### Iterative Enforcement and CI Completion

- [ ] **FLOW-01**: Enforcement workflow includes iterative loop (run gate, detect violations, apply minimal refactor, re-run) until full pass.
- [ ] **FLOW-02**: Implementation evidence includes at least one failing canonical run and one passing canonical run.
- [ ] **FLOW-03**: Project is considered complete only when all gate checks pass in a single canonical execution.
- [ ] **FLOW-04**: CI marks canonical quality job as required for merge.
- [ ] **FLOW-05**: Documentation reflects enforcement policy, non-bypass rule, and completion definition.

## v2 Requirements

None.

## Out of Scope

| Feature | Reason |
|---------|--------|
| New end-user malware-analysis features unrelated to enforcement | Not required to satisfy deterministic quality gate objective |
| Relaxing thresholds or introducing temporary ignores to get green CI | Violates strict hard-gate policy |
| Supporting multiple equivalent quality entrypoints | Violates single canonical gate contract |

## Traceability

Which phases cover which requirements. Updated during roadmap creation.

| Requirement | Phase | Status |
|-------------|-------|--------|
| GATE-01 | Phase 1 | Complete |
| GATE-02 | Phase 1 | Complete |
| GATE-03 | Phase 2 | Complete |
| GATE-04 | Phase 2 | Complete |
| GATE-05 | Phase 2 | Complete |
| GATE-06 | Phase 1 | Complete |
| CC-01 | Phase 3 | Complete |
| CC-02 | Phase 3 | Complete |
| CC-03 | Phase 3 | Complete |
| CC-04 | Phase 4 | In Progress |
| CC-05 | Phase 5 | Pending |
| CC-06 | Phase 5 | Pending |
| CC-07 | Phase 5 | Pending |
| CC-08 | Phase 5 | Pending |
| CC-09 | Phase 5 | Pending |
| CC-10 | Phase 6 | Pending |
| CC-11 | Phase 6 | Pending |
| CC-12 | Phase 6 | Pending |
| CC-13 | Phase 6 | Pending |
| ARCH-01 | Phase 7 | Pending |
| ARCH-02 | Phase 8 | Pending |
| ARCH-03 | Phase 8 | Pending |
| ARCH-04 | Phase 8 | Pending |
| ARCH-05 | Phase 8 | Pending |
| ARCH-06 | Phase 7 | Pending |
| ARCH-07 | Phase 7 | Pending |
| TEST-01 | Phase 4 | In Progress |
| TEST-02 | Phase 4 | In Progress |
| TEST-03 | Phase 3 | Complete |
| FLOW-01 | Phase 9 | Pending |
| FLOW-02 | Phase 9 | Pending |
| FLOW-03 | Phase 9 | Pending |
| FLOW-04 | Phase 2 | Complete |
| FLOW-05 | Phase 9 | Pending |

**Coverage:**
- v1 requirements: 34 total
- Mapped to phases: 34
- Unmapped: 0

---
*Requirements defined: 2026-02-28*
*Last updated: 2026-02-28 after FLOW-04 enforcement on main*
