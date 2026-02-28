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

# Rulesets inspection/apply path
gh api repos/seifreed/docir/rulesets

# Branch protection inspection/apply path
gh api repos/seifreed/docir/branches/main/protection --include
```

## Verification Contract

A merge to `main` is compliant with `FLOW-04` only when required-check configuration includes `quality-gate` as an active required status context.

## Evidence (2026-02-28, Gap-Closure Plan 02-04)

Repository and default branch:

```text
seifreed/docir main
```

Authentication status summary:

```text
github.com account: seifreed
scopes: gist, read:org, repo, workflow
```

Rulesets API result:

```text
gh api repos/seifreed/docir/rulesets
HTTP 403 Forbidden
{"message":"Upgrade to GitHub Pro or make this repository public to enable this feature.","documentation_url":"https://docs.github.com/rest/repos/rules#get-all-repository-rulesets","status":"403"}
exit:1
```

Branch protection API result:

```text
gh api repos/seifreed/docir/branches/main/protection --include
HTTP 403 Forbidden
{"message":"Upgrade to GitHub Pro or make this repository public to enable this feature.","documentation_url":"https://docs.github.com/rest/branches/branch-protection#get-branch-protection","status":"403"}
exit:1
```

Status on 2026-02-28:

- Branch protection API exit code: 1
- Rulesets API exit code: 1
- `GATE-05` is complete (`quality-gate` workflow/job exists and runs canonical gate).
- `FLOW-04` remains blocked by repository tier limits until branch-protection/ruleset features are available.

## Blocker Outcome

`FLOW-04` cannot be marked complete in this repository state because required-check enforcement endpoints are gated (`HTTP 403`) despite valid authenticated access.

Unblock options:

1. Upgrade account/repository tier to enable private-repo branch protection/rulesets.
2. Make repository public (if policy allows), then configure `quality-gate` as required.
