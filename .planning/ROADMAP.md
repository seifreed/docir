# Roadmap: docir Deterministic Quality Enforcement

## Overview

This roadmap delivers deterministic, non-bypassable quality enforcement for `docir` through one canonical gate, hard policy checks, architecture boundary enforcement, and completion evidence that proves a full pass in one canonical execution.

## Phases

**Phase Numbering:**
- Integer phases (1, 2, 3): Planned milestone work
- Decimal phases (2.1, 2.2): Urgent insertions (marked with INSERTED)

Decimal phases appear between their surrounding integers in numeric order.

- [ ] **Phase 1: Canonical Gate Surface** - Establish the single accepted quality entrypoint and strict pass/fail contract.
- [ ] **Phase 2: Workflow Routing** - Route local, pre-commit, and CI execution through the canonical gate.
- [ ] **Phase 3: Baseline Clean Code Commands** - Enforce formatting, linting, tests, and warning posture in canonical runs.
- [ ] **Phase 4: Coverage Integrity Enforcement** - Enforce >=95% coverage and anti-inflation test integrity in canonical runs.
- [ ] **Phase 5: Forbidden Construct Policy** - Block prohibited runtime constructs through hard gate failures.
- [ ] **Phase 6: Code Hygiene and Complexity Policy** - Enforce dead-code, imports, documentation, and complexity thresholds.
- [ ] **Phase 7: Architecture Policy Definition** - Define enforceable layer model and dependency-inversion boundaries.
- [ ] **Phase 8: Architecture Violation Enforcement** - Fail canonical runs on forbidden dependencies and circular coupling.
- [ ] **Phase 9: Iterative Enforcement Evidence and Completion Contract** - Prove failing-to-passing loop and codify project completion rules.

## Phase Details

### Phase 1: Canonical Gate Surface
**Goal**: Users and CI have exactly one valid quality gate entrypoint with deterministic exit behavior.
**Depends on**: Nothing (first phase)
**Requirements**: GATE-01, GATE-02, GATE-06
**Success Criteria** (what must be TRUE):
  1. Running `./scripts/quality_gate.sh` is the only accepted quality-gate invocation path in the repository.
  2. A fully clean repository run returns exit code `0`, and any failed check returns a non-zero code.
  3. No alternate script can be used as a substitute quality gate entrypoint.
**Plans**: TBD

### Phase 2: Workflow Routing
**Goal**: All routine quality workflows consistently execute the canonical gate.
**Depends on**: Phase 1
**Requirements**: GATE-03, GATE-04, GATE-05, FLOW-04
**Success Criteria** (what must be TRUE):
  1. Developers can follow documented local workflow steps that call only the canonical gate.
  2. Pre-commit workflow documentation points to the canonical gate and no parallel quality path.
  3. CI required checks execute `./scripts/quality_gate.sh` directly as the merge-blocking job.
**Plans**: TBD

### Phase 3: Baseline Clean Code Commands
**Goal**: Canonical runs always execute baseline formatting, linting, testing, and warning-strict checks.
**Depends on**: Phase 2
**Requirements**: CC-01, CC-02, CC-03, TEST-03
**Success Criteria** (what must be TRUE):
  1. Canonical gate fails when formatting drift exists and passes only when formatting is compliant.
  2. Canonical gate fails on Clippy warnings because warnings are denied.
  3. Canonical gate fails when tests fail and passes when the test suite passes.
  4. No warning suppression path is used to force a pass.
**Plans**: TBD

### Phase 4: Coverage Integrity Enforcement
**Goal**: Coverage threshold and test integrity are enforced as non-optional gate requirements.
**Depends on**: Phase 3
**Requirements**: CC-04, TEST-01, TEST-02
**Success Criteria** (what must be TRUE):
  1. Canonical gate reports and enforces coverage >=95% using `cargo llvm-cov`.
  2. CI execution cannot skip coverage measurement in the canonical run.
  3. Coverage-passing tests demonstrate meaningful behavior validation rather than trivial inflation patterns.
**Plans**: TBD

### Phase 5: Forbidden Construct Policy
**Goal**: Canonical runs reject prohibited constructs that undermine reliability and maintainability.
**Depends on**: Phase 4
**Requirements**: CC-05, CC-06, CC-07, CC-08, CC-09
**Success Criteria** (what must be TRUE):
  1. Canonical gate fails when `unwrap()` appears in non-test code.
  2. Canonical gate fails when `expect()` appears in non-test code.
  3. Canonical gate fails when `panic!()` appears in production code.
  4. Canonical gate fails when `todo!()` or `unimplemented!()` appears in repository code.
**Plans**: TBD

### Phase 6: Code Hygiene and Complexity Policy
**Goal**: Canonical runs enforce hygiene and maintainability constraints for codebase health.
**Depends on**: Phase 5
**Requirements**: CC-10, CC-11, CC-12, CC-13
**Success Criteria** (what must be TRUE):
  1. Canonical gate fails when dead code exists.
  2. Canonical gate fails when unused imports are present.
  3. Canonical gate fails when a public function is missing documentation.
  4. Canonical gate fails when any function exceeds cyclomatic complexity 10.
**Plans**: TBD

### Phase 7: Architecture Policy Definition
**Goal**: The repository has explicit, enforceable architecture boundaries and inversion contracts.
**Depends on**: Phase 6
**Requirements**: ARCH-01, ARCH-06, ARCH-07
**Success Criteria** (what must be TRUE):
  1. Architecture policy clearly defines `domain`, `application`, `infrastructure`, and `presentation` layers for workspace components.
  2. Cross-layer contracts use trait-based dependency inversion where boundaries are crossed.
  3. Domain layer policy explicitly forbids external framework leakage and is testable by the gate.
**Plans**: TBD

### Phase 8: Architecture Violation Enforcement
**Goal**: Canonical runs fail deterministically on forbidden dependency directions and circular coupling.
**Depends on**: Phase 7
**Requirements**: ARCH-02, ARCH-03, ARCH-04, ARCH-05
**Success Criteria** (what must be TRUE):
  1. Canonical gate fails if domain depends on infrastructure.
  2. Canonical gate fails if domain or application depends on presentation.
  3. Canonical gate fails when circular crate dependencies are introduced.
**Plans**: TBD

### Phase 9: Iterative Enforcement Evidence and Completion Contract
**Goal**: Enforcement loop evidence and completion definition are documented and demonstrable.
**Depends on**: Phase 8
**Requirements**: FLOW-01, FLOW-02, FLOW-03, FLOW-05
**Success Criteria** (what must be TRUE):
  1. Project evidence shows iterative enforcement loop usage from failure detection to minimal fix and re-run.
  2. Evidence includes at least one failing canonical run and one passing canonical run.
  3. Documentation states the project is complete only when all checks pass in one canonical execution.
  4. Documentation codifies non-bypass enforcement policy and definition of done.
**Plans**: TBD

## Progress

**Execution Order:**
Phases execute in numeric order: 1 -> 2 -> 3 -> 4 -> 5 -> 6 -> 7 -> 8 -> 9

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. Canonical Gate Surface | 0/TBD | Not started | - |
| 2. Workflow Routing | 0/TBD | Not started | - |
| 3. Baseline Clean Code Commands | 0/TBD | Not started | - |
| 4. Coverage Integrity Enforcement | 0/TBD | Not started | - |
| 5. Forbidden Construct Policy | 0/TBD | Not started | - |
| 6. Code Hygiene and Complexity Policy | 0/TBD | Not started | - |
| 7. Architecture Policy Definition | 0/TBD | Not started | - |
| 8. Architecture Violation Enforcement | 0/TBD | Not started | - |
| 9. Iterative Enforcement Evidence and Completion Contract | 0/TBD | Not started | - |
