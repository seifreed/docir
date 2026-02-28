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
gh repo view --json nameWithOwner,defaultBranchRef,isPrivate

# Apply branch protection with required status check
gh api -X PUT repos/seifreed/docir/branches/main/protection \
  -H "Accept: application/vnd.github+json" \
  --input /tmp/protection_payload.json

# Verify active protection
gh api repos/seifreed/docir/branches/main/protection --jq '{strict:.required_status_checks.strict, checks:.required_status_checks.checks}'
```

## Verification Contract

A merge to `main` is compliant with `FLOW-04` only when required-check configuration includes `quality-gate` as an active required status context.

## Evidence (2026-02-28, Gap-Closure Plan 02-04)

Repository state:

```text
repo: seifreed/docir
default branch: main
visibility: PUBLIC
```

Applied protection payload (summary):

```json
{
  "required_status_checks": {
    "strict": true,
    "contexts": ["quality-gate"]
  },
  "enforce_admins": false,
  "required_pull_request_reviews": null,
  "restrictions": null
}
```

Verification output:

```json
{
  "strict": true,
  "checks": [
    {
      "context": "quality-gate",
      "app_id": null
    }
  ]
}
```

Status on 2026-02-28:

- Branch protection endpoint reachable and configured.
- Required status check context `quality-gate` active on `main`.
- `GATE-05` complete.
- `FLOW-04` now enforceable and evidence-backed.

## Outcome

`FLOW-04` is satisfied: canonical CI job `quality-gate` is configured as a required merge check for `main`.
