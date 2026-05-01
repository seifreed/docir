#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "${repo_root}"

python3 - <<'PY'
import sys
from pathlib import Path

BASELINE_FILE = Path("scripts/quality_phase1_baseline.sh")
LAYER_FILE = Path("scripts/quality_layer_policy.sh")
ALLOWLIST_FILE = Path("scripts/lib/dependency_allowlist.sh")
REQUIRED_SOURCE_LINE = 'source "${REPO_ROOT}/scripts/lib/dependency_allowlist.sh"'

if not ALLOWLIST_FILE.exists():
    print("quality_policy_consistency: FAIL")
    print("  missing allowlist source file: scripts/lib/dependency_allowlist.sh")
    sys.exit(1)

errors = []
for script in (BASELINE_FILE, LAYER_FILE):
    text = script.read_text(encoding="utf-8", errors="replace")
    if REQUIRED_SOURCE_LINE not in text:
        errors.append(f"{script}: missing shared allowlist source line")
    if "is_allowed_dependency()" in text:
        errors.append(f"{script}: local is_allowed_dependency() should not exist")

allowlist_text = ALLOWLIST_FILE.read_text(encoding="utf-8", errors="replace")
if "is_allowed_dependency()" not in allowlist_text:
    errors.append("scripts/lib/dependency_allowlist.sh: missing is_allowed_dependency()")

if errors:
    print("quality_policy_consistency: FAIL")
    for error in errors:
        print(f"  {error}")
    sys.exit(1)

print("quality_policy_consistency: PASS")
PY
