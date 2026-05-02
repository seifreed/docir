# Roadmap: docir Deterministic Quality Enforcement

## Overview

This roadmap delivers deterministic, non-bypassable quality enforcement for `docir` through one canonical gate, hard policy checks, architecture boundary enforcement, and completion evidence that proves a full pass in one canonical execution.

## Phases

**Phase Numbering:**
- Integer phases (1, 2, 3): Planned milestone work
- Decimal phases (2.1, 2.2): Urgent insertions (marked with INSERTED)

Decimal phases appear between their surrounding integers in numeric order.

- [x] **Phase 1: Canonical Gate Surface** - Establish the single accepted quality entrypoint and strict pass/fail contract.
- [x] **Phase 2: Workflow Routing** - Route local, pre-commit, and CI execution through the canonical gate.
- [x] **Phase 3: Baseline Clean Code Commands** - Enforce formatting, linting, tests, and warning posture in canonical runs.
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
**Plans**:
- [x] `01-01-PLAN.md` - Canonical gate entrypoint scaffold and contract smoke test
- [x] `01-02-PLAN.md` - Exit code class semantics and scenario verification
- [x] `01-03-PLAN.md` - Canonical-only policy documentation and non-bypass inventory

### Phase 2: Workflow Routing
**Goal**: All routine quality workflows consistently execute the canonical gate.
**Depends on**: Phase 1
**Requirements**: GATE-03, GATE-04, GATE-05, FLOW-04
**Success Criteria** (what must be TRUE):
  1. Developers can follow documented local workflow steps that call only the canonical gate.
  2. Pre-commit workflow documentation points to the canonical gate and no parallel quality path.
  3. CI required checks execute `./scripts/quality_gate.sh` directly as the merge-blocking job.
**Plans**:
- [x] `02-01-PLAN.md` - Local workflow documentation routed to canonical gate
- [x] `02-02-PLAN.md` - Pre-commit workflow routing and hook installer alignment
- [x] `02-03-PLAN.md` - Canonical CI job plus required-check runbook baseline
- [x] `02-04-PLAN.md` - Gap-closure with live GitHub API verification and required-check enforcement

### Phase 3: Baseline Clean Code Commands
**Goal**: Canonical runs always execute baseline formatting, linting, testing, and warning-strict checks.
**Depends on**: Phase 2
**Requirements**: CC-01, CC-02, CC-03, TEST-03
**Success Criteria** (what must be TRUE):
  1. Canonical gate fails when formatting drift exists and passes only when formatting is compliant.
  2. Canonical gate fails on Clippy warnings because warnings are denied.
  3. Canonical gate fails when tests fail and passes when the test suite passes.
  4. No warning suppression path is used to force a pass.
**Plans**:
- [x] `03-01-PLAN.md` - Baseline command execution in canonical gate with deterministic failure classification.
- [x] `03-02-PLAN.md` - Baseline command contract harness and strict warning posture documentation.

### Phase 4: Coverage Integrity Enforcement
**Goal**: Coverage threshold and test integrity are enforced as non-optional gate requirements.
**Depends on**: Phase 3
**Requirements**: CC-04, TEST-01, TEST-02
**Success Criteria** (what must be TRUE):
  1. Canonical gate reports and enforces coverage >=95% using `cargo llvm-cov`.
  2. CI execution cannot skip coverage measurement in the canonical run.
  3. Coverage-passing tests demonstrate meaningful behavior validation rather than trivial inflation patterns.
**Plans**:
- [x] `04-01-PLAN.md` - Canonical coverage stage, threshold contract, and CI non-skip enforcement.
- [x] `04-02-PLAN.md` - Coverage integrity behavior tests and anti-inflation policy alignment.
- [x] `04-03-PLAN.md` - Gap closure with targeted parser/security hotspot behavior tests.
- [x] `04-04-PLAN.md` - Additional hotspot behavior validation and canonical re-measure.
- [x] `04-05-PLAN.md` - ODF hotspot behavior tests and refreshed residual coverage inventory.
- [x] `04-06-PLAN.md` - OOXML/RTF hotspot behavior tests and canonical residual refresh.
- [x] `04-07-PLAN.md` - ODF residual hotspot closure with canonical re-measure and next shortlist.
- [x] `04-08-PLAN.md` - Residual hotspot closure continuation with canonical re-measure and updated shortlist.
- [x] `04-09-PLAN.md` - Residual hotspot closure increment with canonical re-measure and bounded handoff.
- [x] `04-10-PLAN.md` - Bounded closure increment requiring canonical fail-under 95 pass for acceptance.
- [x] `04-11-PLAN.md` - Residual closure continuation with canonical re-measure and deterministic handoff.
- [x] `04-12-PLAN.md` - Residual closure continuation with canonical re-measure and deterministic handoff.
- [x] `04-13-PLAN.md` - Residual closure continuation with canonical re-measure and deterministic handoff.
- [x] `04-14-PLAN.md` - Cross-crate hotspot closure increment with canonical re-measure and blocker ranking.
- [x] `04-15-PLAN.md` - Cross-crate hotspot closure in docir-diff/docir-rules with canonical re-measure.
- [x] `04-16-PLAN.md` - Parser hotspot closure in hwp legacy helpers with canonical re-measure.
- [x] `04-17-PLAN.md` - Cross-crate index hotspot closure with canonical re-measure and refreshed ranking.
- [x] `04-18-PLAN.md` - Utility hotspot closure in diff/core/parser/rules with canonical re-measure and blocker logging.
- [x] `04-19-PLAN.md` - Parser `rtf/core.rs` hotspot closure with canonical re-measure and refreshed ranking.

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

## Fase 8 de cierre (Plan del cierre operativo solicitado)

### Objetivo

Aplicar una re-auditoría completa del estado actual con el mismo esquema de evaluación, registrar los resultados contra metas objetivo `>=8/10`, y actualizar este roadmap con la deuda remanente concreta.

### Resultado de cierre (2026-03-01)

- **Clean Code:** `6/10` (bloqueantes: 8 ficheros >800 LOC, 11 funciones >100 LOC, 86 usos de `unwrap/expect/panic/unreachable` en producción).
- **Clean Architecture:** `7/10` (sin nuevas violaciones declarativas por dependencia, pero sin enforcement de prohibiciones de capa/ciclos aún implementado en CI).
- **Simplification:** `5/10` (persisten módulos de gran tamaño, lógicas de control repetidas y deudas de control en scripts).
- **Metas objetivo actualizadas:** `Clean Code >= 8`, `Clean Architecture >= 8`, `Simplification >= 8` como criterio de seguimiento de cierre.

### Ruta de roadmap actualizada (remanente)

1. **Semana 1 (bloqueo inmediato):**
   - Reducir archivos críticos: `crates/docir-diff/src/summary.rs`, `crates/docir-parser/src/ooxml/xlsx/styles.rs`, `crates/docir-parser/src/ooxml/docx/document/table.rs`, `crates/docir-parser/src/ooxml/xlsx/parser/tests.rs`, `crates/docir-parser/src/ooxml/docx/document/tests.rs`, `crates/docir-parser/src/ooxml/pptx/tests.rs`, `crates/docir-parser/src/odf/ods_tests.rs`, `crates/docir-parser/src/odf/presentation_helpers.rs`.
   - Reducir 86 llamadas productivas de constructores prohibidos (`unwrap/expect/panic/unreachable`) por capa, priorizando parser/core/security.
   - Armonizar scripts de control para evitar fallos de semántica de shell (`mapfile`/compatibilidad), manteniendo `quality_gate.sh` como único contrato.

2. **Semana 2:**
   - Activar CC-05/CC-06/CC-07 con verificación de cambios y estado de salida de gate.
   - Completar fases 5/6/7 del roadmap principal con controles de higiene, complejidad y límites de arquitectura.
   - Documentar política de capa (domain/app/infra/presentation) y reglas de no-fugas.

3. **Semana 3:**
   - Añadir evidencia de ciclo de mejora (fallo -> arreglo mínimo -> re-ejecución del gate).
   - Ejecutar medición final y actualizar este roadmap solo si los tres scores alcanzan `>=8/10`.

## Progress

**Execution Order:**
Phases execute in numeric order: 1 -> 2 -> 3 -> 4 -> 5 -> 6 -> 7 -> 8 -> 9

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. Canonical Gate Surface | 3/3 | Complete | 2026-02-28 |
| 2. Workflow Routing | 4/4 | Complete | 2026-02-28 |
| 3. Baseline Clean Code Commands | 2/2 | Complete | 2026-02-28 |
| 4. Coverage Integrity Enforcement | 19/19 | In progress (executed; canonical threshold unmet) | - |
| 5. Forbidden Construct Policy | 0/TBD | Not started | - |
| 6. Code Hygiene and Complexity Policy | 0/TBD | In progress | - |
| 7. Architecture Policy Definition | 0/TBD | Not started | - |
| 8. Architecture Violation Enforcement | 0/TBD | Not started | - |
| 9. Iterative Enforcement Evidence and Completion Contract | 0/TBD | In progress | - |

### Phase 9: evidencias y bloqueos actuales (actualización)

- Plan de cierre activo: ejecutar un ciclo de mejora en dos pasos con evidencia persistente:
  - **Step A:** `./scripts/quality_gate.sh` -> fallo registrado y capturado en `/tmp/quality_gate_pre.log`.
  - **Step B:** corrección de heurísticas CC-12/CC-13 + rerun registrado en `/tmp/quality_gate_wrapped.log`.
- Resultado tras el ciclo:
  - `CC-12 count` bajó de `495` a `184`.
  - `CC-13 count` bajó de `15` a `0`.
- Bloqueo residual para cierre de fase:
  - CC-12 (documentación pública) permanece en `184` y no está cerrado en esta iteración.
- Evidencia persistente: `docs/quality-phase-8-iteration-loop-evidence.md`.
