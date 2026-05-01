#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "${repo_root}"

if [ ! -x "./scripts/quality_phase1_baseline.sh" ]; then
  echo "Missing executable baseline script: ./scripts/quality_phase1_baseline.sh"
  exit 1
fi

baseline_log="$(mktemp)"
inventory_log="$(mktemp)"
trap 'rm -f "${baseline_log}" "${inventory_log}"' EXIT

bash ./scripts/quality_phase1_baseline.sh >"${baseline_log}"

baseline_report="$(awk '/^Output:/ {print $2}' "${baseline_log}" | tail -n 1)"
if [ -z "${baseline_report}" ] || [ ! -f "${baseline_report}" ]; then
  echo "Baseline report path missing or not found."
  cat "${baseline_log}"
  exit 1
fi

required_labels=(
  "Rust files (src)"
  "Production files (src)"
  "Total LOC"
  "Files over 800 LOC"
  "Functions over 100 LOC (heuristic)"
  "Panic/unwrap/expect calls in production"
  "Architecture dependency violations"
)

for label in "${required_labels[@]}"; do
  if ! grep -Fq "| ${label} |" "${baseline_report}"; then
    echo "Missing baseline metric label: ${label}"
    exit 1
  fi
done

bash ./scripts/quality_no_unwrap_expect_in_production.sh inventory >"${inventory_log}"
inventory_total="$(awk -F= '/^Summary: total=/ {print $2}' "${inventory_log}" | tr -d '[:space:]')"
baseline_total="$(
  python3 - "${baseline_report}" <<'PY'
import re
import sys
text = open(sys.argv[1], encoding="utf-8").read()
m = re.search(r"\|\s*Panic/unwrap/expect calls in production\s*\|\s*([0-9]+)\s*\|", text)
print(m.group(1) if m else "")
PY
)"

if [ -z "${baseline_total}" ]; then
  echo "Unable to extract panic/unwrap/expect metric from baseline report."
  exit 1
fi

if [ "${baseline_total}" != "${inventory_total}" ]; then
  echo "Baseline/inventory mismatch: baseline=${baseline_total}, inventory=${inventory_total}"
  exit 1
fi

echo "quality_phase1_baseline_contract: OK"
