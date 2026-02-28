# Pre-Commit Quality Workflow

## Purpose

This repository uses a versioned Git pre-commit hook that delegates quality acceptance to the canonical gate only.

## Setup

From repository root:

```bash
./scripts/install_hooks.sh
```

Expected output includes:

- `Configured core.hooksPath=.githooks`

## Contract

- Git loads `.githooks/pre-commit` via `core.hooksPath`.
- The hook resolves repository root and `exec`s `./scripts/quality_gate.sh`.
- Commit acceptance authority remains the canonical gate only.
- Hook pass/fail behavior mirrors the gate exit code.
- Hook output preserves the final `QUALITY_GATE_RESULT=...` line emitted by the canonical gate.

## Verification

```bash
# Confirm deterministic hook path
./scripts/install_hooks.sh
git config --get core.hooksPath

# Confirm hook delegates to canonical gate and propagates failure
QUALITY_GATE_FORCE_FAIL=1 ./.githooks/pre-commit

# Confirm canonical exit-code behavior contract remains valid
bash scripts/tests/quality_gate_exit_codes.sh
```

Expected results:

- `git config --get core.hooksPath` returns `.githooks`.
- Forced failure exits non-zero and prints final `QUALITY_GATE_RESULT=FAIL ...` line.
- Exit-code contract test reports `quality_gate_exit_codes: OK`.
