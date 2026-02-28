# CI Required Quality Check Runbook

## Canonical Check Name

- Required check context: `quality-gate`
- Source: `.github/workflows/quality-gate.yml` job `quality-gate`
- Acceptance command in CI: `./scripts/quality_gate.sh`

## GitHub UI Path

1. Open repository settings.
2. Go to `Settings -> Branches` (branch protection) or `Settings -> Rules -> Rulesets`.
3. Select target branch policy for `main`.
4. Enable required status checks.
5. Add required check context `quality-gate`.
6. Save and verify merge is blocked when `quality-gate` is failing or absent.

## GitHub CLI/API Path

```bash
# Repository and default branch
gh repo view --json nameWithOwner,defaultBranchRef

# Branch protection check
gh api repos/seifreed/docir/branches/main/protection --include

# Rulesets check
gh api repos/seifreed/docir/rulesets
```

## Verification Contract

A merge to `main` is compliant with `FLOW-04` only when required-check configuration includes `quality-gate` as an active required status context.
