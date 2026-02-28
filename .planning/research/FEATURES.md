# Feature Research

**Domain:** Deterministic quality-gate enforcement for Rust workspaces
**Researched:** 2026-02-28
**Confidence:** HIGH

## Feature Landscape

### Table Stakes (Users Expect These)

Features users assume exist. Missing these = product feels incomplete.

| Feature | Why Expected | Complexity | Notes |
|---------|--------------|------------|-------|
| Single canonical gate entrypoint (`./scripts/quality_gate.sh`) | Teams expect one command for local + CI parity | MEDIUM | Dependency: CI job must call only this script; aligns with non-bypass requirement in `.planning/PROJECT.md`. |
| Deterministic gate ordering and fail-fast behavior | Developers expect reproducible outcomes and clear first failure | MEDIUM | Dependency: stable check sequence (`fmt -> lint -> test -> coverage -> policy`). |
| Hard policy enforcement for banned patterns | Quality-gate systems are expected to block obvious reliability hazards | MEDIUM | Dependency: static scanners for `unwrap/expect/panic/todo/unimplemented`, dead code, unused imports, and doc coverage policy. |
| Layering/dependency rule checks (architecture gate) | Architecture drift control is baseline in serious multi-crate repos | HIGH | Dependency: explicit crate/layer map and rule engine for allowed dependency direction. |
| CI required-status integration (merge blocking) | Gate systems are expected to enforce quality before merge | LOW | Dependency: branch protection + required check naming stability. |
| Coverage threshold enforcement (>=95%) | Users expect quantitative regression protection | MEDIUM | Dependency: `cargo llvm-cov` availability and stable exclusions policy. |
| Machine-readable gate report artifacts | Teams expect auditable outputs for failures and trends | MEDIUM | Dependency: structured output (JSON/markdown summary) and CI artifact upload path. |

### Differentiators (Competitive Advantage)

Features that set the product apart. Not required, but valuable.

| Feature | Value Proposition | Complexity | Notes |
|---------|-------------------|------------|-------|
| Violation taxonomy tied to risk domains (memory, silent parse failures, panic surfaces, architecture) | Converts generic lint failures into operationally meaningful risk signals | HIGH | Dependency: mapping rules to `CONCERNS.md` categories and ownership metadata. |
| Drift detection for repeated parser control-flow patterns | Prevents behavior divergence across OOXML/RTF/ODF parser loops identified in concerns | HIGH | Dependency: AST/pattern checker plus baseline signatures for loop/match motifs. |
| Security extraction consistency gate across DOCX/XLSX/PPTX | Enforces uniform threat-signal extraction across formats | HIGH | Dependency: shared relationship classification contract + cross-format fixture matrix. |
| Determinism guardrails for output identity contracts | Flags non-deterministic identifiers that can break reproducibility workflows | MEDIUM | Dependency: export-layer normalization checks and reproducibility mode tests. |
| Guided remediation hints with precise file/module pointers | Reduces time-to-fix for failed gates without weakening strictness | MEDIUM | Dependency: rule metadata with canonical guidance and source mapping. |

### Anti-Features (Commonly Requested, Often Problematic)

Features that seem good but create problems.

| Feature | Why Requested | Why Problematic | Alternative |
|---------|---------------|-----------------|-------------|
| Multiple gate scripts per team/language area | Teams want autonomy and faster local iteration | Creates bypass paths and inconsistent enforcement outcomes | Keep one canonical script; allow scoped modes via strict flags inside same entrypoint. |
| Warning-only mode for policy violations in mainline CI | Reduces short-term friction while debt exists | Normalizes non-compliance and defeats deterministic enforcement goal | Use hard-fail in required CI; optionally allow advisory mode only on non-protected branches. |
| Auto-suppressions for flaky checks | Appears to stabilize CI quickly | Hides root causes and accumulates silent quality debt | Track flakiness explicitly, quarantine by issue ID + owner + expiry, still block on protected branches when expired. |
| Per-crate custom thresholds without global floor | Teams want flexibility for legacy crates | Produces uneven quality bars and cross-crate weak links | Maintain global minimums with temporary, time-boxed exceptions documented in policy. |
| Expanding gate to runtime performance benchmarking by default | Sounds comprehensive | Increases nondeterminism and runtime variance; can destabilize merge velocity | Keep performance checks separate as scheduled/perf pipeline with stable baselines. |

## Feature Dependencies

```
[Canonical gate entrypoint]
    └──requires──> [Deterministic check ordering]
                       └──requires──> [Toolchain/version pinning]

[Architecture dependency checks]
    └──requires──> [Layer map + allowed dependency matrix]

[Coverage threshold enforcement]
    └──requires──> [Stable coverage toolchain + exclusions policy]

[Risk taxonomy reporting] ──enhances──> [Hard policy enforcement]
[Remediation hints] ──enhances──> [Developer fix velocity]

[Multiple gate scripts] ──conflicts──> [Single canonical gate contract]
[Warning-only CI mode] ──conflicts──> [Non-bypass deterministic enforcement]
```

### Dependency Notes

- **Canonical gate entrypoint requires deterministic check ordering:** one script is only useful if all environments execute the same sequence and semantics.
- **Architecture dependency checks require a layer map + rules:** without explicit allowed edges, violations cannot be evaluated objectively.
- **Coverage threshold enforcement requires stable tooling:** threshold disputes and noise increase without pinned coverage collection behavior.
- **Risk taxonomy reporting enhances hard policy enforcement:** enriches failures with actionable context while preserving hard-fail semantics.
- **Remediation hints enhance developer fix velocity:** keeps strict gate while reducing turnaround time for high-volume violations.
- **Multiple gate scripts conflict with canonical contract:** parallel entrypoints reintroduce bypass and drift risk.
- **Warning-only CI mode conflicts with deterministic enforcement:** non-blocking policies make quality optional on protected branches.

## MVP Definition

### Launch With (v1)

Minimum viable product — what's needed to validate the concept.

- [ ] Canonical gate script as single entrypoint — essential to eliminate alternate paths.
- [ ] Deterministic ordered checks with fail-fast semantics — essential for reproducible outcomes.
- [ ] Hard failures for clean-code policy violations — essential to block known risk patterns.
- [ ] Hard failures for architecture dependency violations — essential to prevent layer drift.
- [ ] Required CI job wired to canonical gate — essential for merge-time enforcement.
- [ ] Coverage threshold check at 95% with documented policy — essential quantitative guardrail.

### Add After Validation (v1.x)

Features to add once core is working.

- [ ] Structured JSON/markdown gate reports — add when baseline enforcement is stable in CI.
- [ ] Risk-domain taxonomy tagging for violations — add once core rule set matures.
- [ ] Guided remediation hints with source pointers — add after recurring failure patterns are observed.
- [ ] Cross-format security extraction consistency checks — add after shared classification contract is established.

### Future Consideration (v2+)

Features to defer until product-market fit is established.

- [ ] Pattern-drift detectors for repeated parser control-flow motifs — defer until false-positive rate is manageable.
- [ ] Determinism audits for identifier/export reproducibility modes — defer until downstream tooling requires strict reproducibility guarantees.
- [ ] Historical trend dashboards and SLOs for gate health — defer until sufficient run history exists.

## Feature Prioritization Matrix

| Feature | User Value | Implementation Cost | Priority |
|---------|------------|---------------------|----------|
| Canonical single gate entrypoint | HIGH | MEDIUM | P1 |
| Deterministic ordered checks | HIGH | MEDIUM | P1 |
| Hard policy enforcement (clean code) | HIGH | MEDIUM | P1 |
| Hard architecture dependency enforcement | HIGH | HIGH | P1 |
| Required CI merge-blocking job | HIGH | LOW | P1 |
| Coverage >=95% enforcement | HIGH | MEDIUM | P1 |
| Structured report artifacts | MEDIUM | MEDIUM | P2 |
| Risk taxonomy mapping | MEDIUM | HIGH | P2 |
| Remediation hint engine | MEDIUM | MEDIUM | P2 |
| Security extraction consistency gate | MEDIUM | HIGH | P2 |
| Parser pattern-drift detection | MEDIUM | HIGH | P3 |
| Determinism reproducibility audit mode | LOW | MEDIUM | P3 |

**Priority key:**
- P1: Must have for launch
- P2: Should have, add when possible
- P3: Nice to have, future consideration

## Competitor Feature Analysis

| Feature | Competitor A | Competitor B | Our Approach |
|---------|--------------|--------------|--------------|
| Canonical gate entrypoint | Often split across make targets/scripts | Often centralized in one pipeline script | Enforce exactly one accepted gate surface (`./scripts/quality_gate.sh`). |
| Policy checks | Usually lint-centric with selective bans | Mixed static analysis with exceptions | Strict ban policy with no bypass path on protected branches. |
| Architecture checks | Frequently absent or informal | Present in some monorepos with custom tooling | Treat layer rules as first-class hard checks for this workspace. |
| Reporting | Basic CI logs | Some machine-readable artifacts | Provide deterministic artifacts and risk-tagged summaries. |
| Coverage enforcement | Common threshold checks, variable rigor | Similar threshold + trend tracking | Fixed global floor (95%) tied to canonical gate and policy docs. |

## Sources

- `.planning/PROJECT.md`
- `.planning/codebase/CONCERNS.md`
- Existing repository constraints (`scripts/quality_gate.sh` as canonical gate surface)

---
*Feature research for: deterministic quality-gate enforcement systems*
*Researched: 2026-02-28*
