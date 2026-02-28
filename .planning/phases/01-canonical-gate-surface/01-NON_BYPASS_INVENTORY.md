# Phase 01 Non-Bypass Inventory

Date: 2026-02-28
Phase: `01-canonical-gate-surface`
Plan: `01-03`

## Scan Scope

- `scripts/` executable and gate-like script names
- Repository documentation surface: `README.md`, `docs/`
- Workflow/config surface: `.github/workflows/`, root workflow helpers
- Planning/policy references that define accepted gate entrypoint

## Scan Commands

```bash
find scripts -maxdepth 3 -type f -print | sort
ls -l scripts scripts/lib scripts/tests
rg -n "quality_gate\\.sh|quality gate|canonical|bypass|cargo fmt|cargo clippy|cargo test" README.md docs scripts .planning
```

## Findings

1. Canonical script exists at `scripts/quality_gate.sh` and is executable.
2. No additional executable gate-like scripts were found under `scripts/`.
3. README quality workflow now documents only `./scripts/quality_gate.sh` as accepted and marks raw check commands as non-authoritative.
4. `docs/quality-gate-policy.md` defines allowed entrypoint, forbidden bypass patterns, and routing expectations.
5. Workflow surface currently has no `.github/workflows/` directory in this repository state, so no CI-side alternate quality invocation is present in-tree yet.

## Policy Conformance

- no alternate accepted gate: confirmed in current repository policy surface.
- Direct raw check commands are documented as non-authoritative for acceptance.
- New workflow/script guidance requires routing through canonical gate.

## Exceptions

None.
